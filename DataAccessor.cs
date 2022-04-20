using OfficeOpenXml;
using ReportExtraction.DTO;
using ReportExtraction.DTO.Enums;
using SharedFunctions.Entities.Sendabox;
using SharedFunctions.Models;
using SharedFunctions.Utilities;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static ReportExtraction.UtilsDataTypes;

namespace ReportExtraction
{
    public class DataAccessor
    {
        private readonly LoggerConfiguration logger = new LoggerConfiguration();
        protected SqlConnection GetSqlConnection()
        {
            return new SqlConnection(ConfigurationManager.ConnectionStrings["sendaboxSql"].ConnectionString);
        }

        public async Task<Queue<ExtractionReportDto>> GetReportParameters()
        {
            try
            {
                using (var connection = GetSqlConnection())
                {
                    var query = @"SELECT
                                    id_extraction_data,
                                    extraction_storedprocedure,
                                    extraction_email,
                                    email_subject,
                                    extraction_format,
                                    file_name,
                                    email_text,
                                    send_compressed,
                                    send_empty,
                                    send_on_business_days,
                                    send_on_week_days,
                                    send_on_month_beginning,
                                    send_on_month_end,
                                    ignore_from_date,
                                    ignore_to_date,
                                    send_by_protocol,
                                    ftp_host,
                                    ftp_port,
                                    ftp_username,
                                    ftp_password,
                                    ftp_remote_folder,
                                    type_of_excel
                                  FROM dbo.EXTRACTION_REPORT
                                  WHERE is_active = 1
                                  ORDER BY id_extraction_data ASC";

                    var cmd = new SqlCommand(query, connection);
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandTimeout = 0;

                    connection.Open();

                    var reader = cmd.ExecuteReader();

                    var paramsQueue = new Queue<ExtractionReportDto>();


                    while (await reader.ReadAsync())
                    {
                        var parameters = new ExtractionReportDto();

                        parameters.ExtractionId = reader["id_extraction_data"].ToString();
                        parameters.StoreProcedure = reader["extraction_storedprocedure"].ToString();
                        parameters.Email = reader["extraction_email"].ToString();
                        parameters.EmailSubject = reader["email_subject"].ToString();

                        if (Enum.TryParse<FileFormatType>(reader["extraction_format"].ToString(), true, out var fileFormat))
                        {
                            parameters.FileFormat = fileFormat;
                        }

                        parameters.FileName = reader["file_name"]?.ToString() ?? parameters.StoreProcedure + ".csv";
                        parameters.EmailText = reader["email_text"].ToString();

                        if (bool.TryParse(reader["send_compressed"].ToString(), out var sendCompressed))
                        {
                            parameters.SendCompressed = sendCompressed;
                        }

                        if (bool.TryParse(reader["send_empty"].ToString(), out var sendEmpty))
                        {
                            parameters.SendEmpty = sendEmpty;
                        }

                        if (bool.TryParse(reader["send_on_business_days"].ToString(), out var sendOnBusinessDays))
                        {
                            parameters.SendOnBusinessDays = sendOnBusinessDays;
                        }

                        parameters.SendOnWeekDays = reader["send_on_week_days"].ToString();

                        if (bool.TryParse(reader["send_on_month_beginning"].ToString(), out var sendOnMonthBeginning))
                        {
                            parameters.SendOnMonthBeginning = sendOnMonthBeginning;
                        }

                        if (bool.TryParse(reader["send_on_month_end"].ToString(), out var sendOnMonthEnd))
                        {
                            parameters.SendOnMonthEnd = sendOnMonthEnd;
                        }

                        var t = reader["ignore_from_date"].ToString();

                        if (reader["ignore_from_date"] != null && DateTime.TryParse(reader["ignore_from_date"].ToString(), out var ignoreDateFrom))
                        {
                            parameters.IgnoreFromDate = ignoreDateFrom;
                        }

                        if (reader["ignore_to_date"] != null && DateTime.TryParse(reader["ignore_to_date"].ToString(), out var ignoreDateTo))
                        {
                            parameters.IgnoreToDate = ignoreDateTo;
                        }

                        if (!string.IsNullOrWhiteSpace(reader["send_by_protocol"].ToString()))
                        {
                            parameters.ProtocolType = (ProtocolType)int.Parse(reader["send_by_protocol"].ToString());
                        }

                        var excelType = reader["type_of_excel"]?.ToString();
                        parameters.TypeOfExcel = string.IsNullOrEmpty(excelType) ? null : (int?)int.Parse(excelType);

                        if (parameters.ProtocolType == ProtocolType.Ftp || parameters.ProtocolType == ProtocolType.MailAndFtp)
                        {
                            parameters.FtpCredentials.FtpHost = reader["ftp_host"].ToString();
                            parameters.FtpCredentials.Port = ushort.Parse(reader["ftp_port"].ToString());
                            parameters.FtpCredentials.Username = reader["ftp_username"].ToString();
                            parameters.FtpCredentials.Password = reader["ftp_password"].ToString();
                            parameters.FtpCredentials.DestinationFolderPath = reader["ftp_remote_folder"].ToString();
                        }

                        paramsQueue.Enqueue(parameters);
                    }

                    return paramsQueue;
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, $"Error on getting export parameters! Exception: [{ex.Message}]", true);
                throw;
            }
        }

        public async Task<ResultAsCsv> CsvFormatBuilder(string storedPrcedure)
        {
            CultureInfo customCulture = (CultureInfo)Thread.CurrentThread.CurrentCulture.Clone();
            customCulture.NumberFormat.NumberDecimalSeparator = ",";

            Thread.CurrentThread.CurrentCulture = customCulture;

            //if (storedPrcedure == "dbo.Export_Daily_Shipments")
            //{
            //    return await BuildCsvWithAddon(storedPrcedure);
            //}
            return await BuildCsv(storedPrcedure);
        }

        public async Task<ResultAsExcel> ExcelFormatBuilder(string storePrcedure, string fileName, int? typeOfExcel = null)
        {
            if (typeOfExcel.HasValue && typeOfExcel == 2)
            {
                return await ExcelMultySheetsFormatBuilder(storePrcedure, fileName);
            }
            else
            {
                return await ExcelSingleSheetFormatBuilder(storePrcedure, fileName);
            }
        }

        public async Task<ResultAsExcel> ExcelSingleSheetFormatBuilder(string storedPrcedure, string fileName)
        {
            using (var connection = GetSqlConnection())
            {
                var cmd = new SqlCommand(storedPrcedure, connection);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 0;

                connection.Open();
                var reader = cmd.ExecuteReader();

                try
                {
                    using (var p = new ExcelPackage())
                    {
                        p.Workbook.Properties.Author = "Sendabox";
                        p.Workbook.Properties.Title = fileName;
                        p.Workbook.Properties.Created = DateTime.Now;

                        p.Workbook.Worksheets.Add(fileName);
                        ExcelWorksheet ws = p.Workbook.Worksheets[1];
                        ws.Name = fileName;
                        ws.Cells.Style.Font.Size = 11;
                        ws.Cells.Style.Font.Name = "Calibri";

                        //Headers 
                        var totalColumns = reader.FieldCount;

                        var columns = Enumerable.Range(0, totalColumns)
                            .Select(reader.GetName).ToArray();

                        for (int col = 0; col < totalColumns; col++)
                        {
                            ws.Cells[1, col + 1].Value = columns[col];
                        }

                        //Data
                        var setString = SetToStringIfDate(reader);

                        var result = new ResultAsExcel();
                        result.HasResults = reader.HasRows ? true : false;

                        int row = 0;
                        while (await reader.ReadAsync())
                        {
                            for (int col = 0; col < totalColumns; col++)
                            {
                                try
                                {
                                    if (setString[col])
                                    {
                                        ws.Cells[row + 2, col + 1].Value = reader.GetValue(col).ToString();
                                    }
                                    else
                                    {
                                        ws.Cells[row + 2, col + 1].Value = reader.GetValue(col);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    logger.LogError(ex, $"Error on building Excel export file at row[{row}] col[{col}]! Stored Procedure: [{storedPrcedure}] {Environment.NewLine}" +
                                        $"Exception: [{ex.Message}]", true);
                                    throw;
                                }
                            }
                            ++row;
                        }

                        ws.Cells[ws.Dimension.Address].AutoFitColumns();
                        for (int col = 1; col <= ws.Dimension.End.Column; col++)
                        {
                            ws.Column(col).Width = ws.Column(col).Width + 1;
                        }
                        result.BytesAsExcel = p.GetAsByteArray();

                        return result;
                    }
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, $"Error on building Excel export file!  Stored Procedure: [{storedPrcedure}] {Environment.NewLine}" +
                                        $"Exception: [{ex.Message}]", true);
                    throw;
                }
            }

        }

        public async Task<ResultAsExcel> ExcelMultySheetsFormatBuilder(string storedPrcedure, string fileName)
        {
            using (var connection = GetSqlConnection())
            {
                var cmd = new SqlCommand(storedPrcedure, connection);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 0;

                connection.Open();
                var reader = cmd.ExecuteReader();

                var sheetNames = new List<string>();

                while (await reader.ReadAsync())
                {
                    sheetNames.Add(reader.GetValue(0).ToString());
                }

                try
                {
                    using (var p = new ExcelPackage())
                    {
                        p.Workbook.Properties.Author = "Sendabox";
                        p.Workbook.Properties.Title = fileName;
                        p.Workbook.Properties.Created = DateTime.Now;

                        var result = new ResultAsExcel();

                        var counter = 0;
                        while (await reader.NextResultAsync())
                        {
                            if (counter >= sheetNames.Count)
                            {
                                var ex = new ArgumentOutOfRangeException();
                                logger.LogError(ex, $"Error on building Excel export file. Sheets names provided do not much sheets in excel documnet!" +
                                    $"Stored Procedure: [{storedPrcedure}]", true);
                                throw ex;
                            }

                            var fullSheetName = counter <= 2 ? sheetNames[counter] : $"{sheetNames[counter]}_{counter + 1}";
                            p.Workbook.Worksheets.Add(fullSheetName);
                            ExcelWorksheet ws = p.Workbook.Worksheets[counter + 1];
                            counter++;

                            ws.Cells.Style.Font.Size = 11;
                            ws.Cells.Style.Font.Name = "Calibri";

                            //Headers 
                            var totalColumns = reader.FieldCount;

                            var columns = Enumerable.Range(0, totalColumns)
                                .Select(reader.GetName).ToArray();

                            for (int col = 0; col < totalColumns; col++)
                            {
                                ws.Cells[1, col + 1].Value = columns[col];
                            }

                            //Data
                            var setString = SetToStringIfDate(reader);

                            result.HasResults = reader.HasRows ? true : false;

                            int row = 0;
                            while (await reader.ReadAsync())
                            {
                                for (int col = 0; col < totalColumns; col++)
                                {
                                    try
                                    {
                                        if (setString[col])
                                        {
                                            ws.Cells[row + 2, col + 1].Value = reader.GetValue(col).ToString();
                                        }
                                        else
                                        {
                                            ws.Cells[row + 2, col + 1].Value = reader.GetValue(col);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        logger.LogError(ex, $"Error on building Excel export file at row[{row}] col[{col}]! Stored Procedure: [{storedPrcedure}] {Environment.NewLine}" +
                                          $"Exception: [{ex.Message}]", true);
                                        throw;
                                    }
                                }
                                ++row;
                            }

                            ws.Cells[ws.Dimension.Address].AutoFitColumns();
                            for (int col = 1; col <= ws.Dimension.End.Column; col++)
                            {
                                ws.Column(col).Width = ws.Column(col).Width + 1;
                            }
                        }
                        if (counter != sheetNames.Count)
                        {
                            var ex = new ArgumentOutOfRangeException();
                            logger.LogError(ex, $"Error on building Excel export file. Sheets names provided do not much sheets in excel documnet! " +
                                $"Stored Procedure: [{storedPrcedure}] {Environment.NewLine}" +
                                          $"Exception: [{ex.Message}]", true);
                            throw ex;
                        }
                        result.BytesAsExcel = p.GetAsByteArray();

                        return result;
                    }
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, $"Error on building Excel export file! Stored Procedure: [{storedPrcedure}] {Environment.NewLine}" +
                                          $"Exception: [{ex.Message}]", true);
                    throw;
                }
            }
        }

        private async Task<ResultAsCsv> BuildCsv(string storePrcedure)
        {
            try
            {
                using (var connection = GetSqlConnection())
                {
                    var cmd = new SqlCommand(storePrcedure, connection);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 0;

                    connection.Open();
                    var reader = cmd.ExecuteReader();

                    var separator = ";";
                    var nL = Environment.NewLine;
                    var dQuote = "\"";

                    var totalColumns = reader.FieldCount;
                    var columns = string.Join(separator,
                        Enumerable.Range(0, totalColumns)
                        .Select(reader.GetName));

                    var sb = new StringBuilder();
                    sb.Append(columns);
                    sb.Append(nL);

                    var addQuote = SetQuotesIfText(reader);

                    var result = new ResultAsCsv();

                    result.HasResults = reader.HasRows ? true : false;

                    while (await reader.ReadAsync())
                    {
                        for (int colIndex = 0; colIndex < totalColumns; colIndex++)
                        {
                            var columnValue = reader
                                .GetValue(colIndex)
                                .ToString()
                                .Trim();

                            if (addQuote[colIndex])
                            {
                                sb.Append(dQuote);
                                sb.Append(columnValue);
                                sb.Append(dQuote);
                            }
                            else
                            {
                                sb.Append(columnValue);
                            }

                            if (colIndex < totalColumns - 1)
                            {
                                sb.Append(separator);
                            }
                        }

                        sb.Append(nL);
                    }
                    result.TextAsCsv = sb.ToString();

                    return result;
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, $"Error on building CSV export file! Stored Procedure: [{storePrcedure}] {Environment.NewLine}" +
                    $"Exception: [{ex.Message}]", true);
                throw;
            }
        }

        private async Task<ResultAsCsv> BuildCsvWithAddon(string storePrcedure)
        {
            try
            {

                var stopWatch = new Stopwatch();

                var resultsDb = new List<DailyExportDbDto>();
                var sbCsv = new StringBuilder();

                var separator = ";";
                var nL = Environment.NewLine;
                var dQuote = "\"";

                var result = new ResultAsCsv();

                using (var connection = GetSqlConnection())
                {
                    var cmd = new SqlCommand(storePrcedure, connection);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 0;

                    connection.Open();
                    var reader = cmd.ExecuteReader();

                    var totalColumns = reader.FieldCount;

                    // IMPORTANT We currently [24/11/2022] take only 24 columns from database for CSV file
                    var currentColumnsUsedForCsv = 24;
                    var columns = string.Join(separator,
                        Enumerable.Range(0, currentColumnsUsedForCsv)
                        .Select(reader.GetName));

                    sbCsv.Append(columns);

                    result.HasResults = reader.HasRows ? true : false;

                    stopWatch.Start();

                    while (await reader.ReadAsync())
                    {
                        var resultDb = new DailyExportDbDto
                        {
                            IdUserCityOrder = ParseToInt(reader["idusercityorder"]),
                            IdUser = ParseToInt(reader["iduser"]),
                            Login = reader["login"].ToString(),
                            OrderPrice = ParseToDouble(reader["orderprice"]),
                            Sales = ParseToDouble(reader["fatturato"]),
                            OrderDate = reader["orderdate"].ToString(),
                            FromProvince = reader["from province"].ToString(),
                            ToCountry = reader["tocountry"].ToString(),
                            ToRegion = reader["toregion"].ToString(),
                            Weight = ParseToFloat(reader["weight"]),
                            WeightRange = reader["weight range"].ToString(),
                            WeightUpperRange = reader["weight upper range"].ToString(),
                            Agent = reader["AGENTE"].ToString(),
                            BuisnessName = reader["businnessname"].ToString(),
                            Status = reader["stato"].ToString(),
                            TargetCountry = reader["target_country"].ToString(),
                            OrderMonth = reader["order_month"].ToString(),
                            OrderWeek = reader["order_week"].ToString(),
                            PaymentMethod = reader["paymentmethod"].ToString(),
                            EstimatedCost = reader["costo stimato"].ToString(),
                            Carrier = reader["carrier"].ToString(),
                            CustomerTracking = reader["customer_tracking"].ToString(),
                            PrivateCompany = reader["azienda_privato"].ToString(),
                            CouponName = reader["CouponName"].ToString(),

                            IdCarrier = ParseToInt(reader["idcarrier"]),
                            FromZip = reader["from_zip"].ToString(),
                            ToZip = reader["to_zip"].ToString(),
                            MultipacketWeight = ParseToDbNullableFloat(reader["multipacket_weight"]),
                            MultipacketDepth = ParseToDbNullableFloat(reader["multipacket_depth"]),
                            MultipacketHeight = ParseToDbNullableFloat(reader["multipacket_height"]),
                            MultipacketWidth = ParseToDbNullableFloat(reader["multipacket_width"]),
                            IdItalianCityFrom = ParseToDbNullableInt(reader["id_italian_city_from"]),
                            IdItalianCityTo = ParseToDbNullableInt(reader["id_italian_city_to"]),
                            IdCountry = ParseToDbNullableInt(reader["id_country"]),
                        };

                        resultsDb.Add(resultDb);
                    }

                }

                var dbAddonsDimensions = AllDbAddonDimensions();
                var dbAddonsBelts = AllDbAddonBelts();
                var dbAddonsMinMaxDim = AllDbAddonMinMaxDim();
                var dbAddonWeight = AllDbAddonWeight();
                var dbAddonsum = AllDbAddonSum();
                var dbAddonPricelist = AllDbAddonPricelist();
                var dbAddonPricelistZipCode = AllDbAddonPricelistZipCode();
                var dbAddonCityZipCode = AllDbAddonCityZipCode();
                var dbCities = AllDbCities();
                var dbAddonForeignCityZipCode = AllDbForeignCityZipCode();
                var dbCountries = AllDbCountries();
                var dbGeonamesEntities = AllDbGeonamesEntity();

                int resultCounter = 0;
                var fileDailyExports = new List<DailyExportFileDto>();

                Parallel.ForEach(resultsDb, resultDb =>
                {
                    var multipacketDimensions = new List<double?>
                                  {
                                      resultDb.MultipacketDepth,
                                      resultDb.MultipacketHeight,
                                      resultDb.MultipacketWidth
                                  };

                    var addon = GetSumOfAddons(
                          resultDb.IdCarrier,
                          resultDb.IdItalianCityFrom,
                          resultDb.IdItalianCityTo,
                          resultDb.FromZip,
                          resultDb.ToZip,
                          resultDb.IdCountry,
                          resultDb.MultipacketWeight,
                          multipacketDimensions,
                          resultDb.OrderPrice,
                          dbAddonsDimensions,
                          dbAddonsBelts,
                          dbAddonsMinMaxDim,
                          dbAddonWeight,
                          dbAddonsum,
                          dbAddonPricelist,
                          dbAddonPricelistZipCode,
                          dbAddonCityZipCode,
                          dbCities,
                          dbAddonForeignCityZipCode,
                          dbCountries,
                          dbGeonamesEntities
                          );

                    var fileDto = MapDailyExportDbToFileDto(resultDb, addon);
                    fileDailyExports.Add(fileDto);

                    resultCounter++;
                });

                var totalTime = stopWatch.Elapsed;

                // add addon as column header
                 sbCsv.Append($"addon{separator}");
                foreach (var de in fileDailyExports)
                {
                    sbCsv.Append(de.IdUserCityOrder);
                    sbCsv.Append(separator);
                    sbCsv.Append(de.IdUser);
                    sbCsv.Append(separator);
                    sbCsv.Append($"{dQuote}{de.Login}{dQuote}");
                    sbCsv.Append(separator);
                    sbCsv.Append(de.OrderPrice);
                    sbCsv.Append(separator);
                    sbCsv.Append(de.Sales);
                    sbCsv.Append(separator);
                    sbCsv.Append(de.OrderDate);
                    sbCsv.Append(separator);
                    sbCsv.Append($"{dQuote}{de.FromProvince}{dQuote}");
                    sbCsv.Append(separator);
                    sbCsv.Append($"{dQuote}{de.ToCountry}{dQuote}");
                    sbCsv.Append(separator);
                    sbCsv.Append($"{dQuote}{de.ToRegion}{dQuote}");
                    sbCsv.Append(separator);
                    sbCsv.Append(de.Weight);
                    sbCsv.Append(separator);
                    sbCsv.Append($"{dQuote}{de.WeightRange}{dQuote}");
                    sbCsv.Append(separator);
                    sbCsv.Append($"{dQuote}{de.WeightUpperRange}{dQuote}");
                    sbCsv.Append(separator);
                    sbCsv.Append($"{dQuote}{de.Agent}{dQuote}");
                    sbCsv.Append(separator);
                    sbCsv.Append($"{dQuote}{de.BuisnessName}{dQuote}");
                    sbCsv.Append(separator);
                    sbCsv.Append($"{dQuote}{de.Status}{dQuote}");
                    sbCsv.Append(separator);
                    sbCsv.Append($"{dQuote}{de.TargetCountry}{dQuote}");
                    sbCsv.Append(separator);
                    sbCsv.Append(de.OrderMonth);
                    sbCsv.Append(separator);
                    sbCsv.Append(de.OrderWeek);
                    sbCsv.Append(separator);
                    sbCsv.Append($"{dQuote}{de.PaymentMethod}{dQuote}");
                    sbCsv.Append(separator);
                    sbCsv.Append(de.EstimatedCost);
                    sbCsv.Append(separator);
                    sbCsv.Append($"{dQuote}{de.Carrier}{dQuote}");
                    sbCsv.Append(separator);
                    sbCsv.Append($"{dQuote}{de.CustomerTracking}{dQuote}");
                    sbCsv.Append(separator);
                    sbCsv.Append($"{dQuote}{de.PrivateCompany}{dQuote}");
                    sbCsv.Append(separator);
                    sbCsv.Append($"{dQuote}{de.CouponName}{dQuote}");
                    sbCsv.Append(separator);
                    sbCsv.Append(de.Addon);
                    sbCsv.Append(separator);

                    sbCsv.Append(nL);
                }

                result.TextAsCsv = sbCsv.ToString();

                return result;
            }
            catch (Exception ex)
            {
                logger.LogError(ex, $"Error on building CSV export file! Stored Procedure: [{storePrcedure}] {Environment.NewLine}" +
                    $"Exception: [{ex.Message}]", true);
                throw;
            }
        }

        public async Task ReadMultipleTables()
        {
            using (var connection = GetSqlConnection())
            {
                var cmd = new SqlCommand("dbo.Export_Utilities_List_Test_Toni", connection);
                cmd.CommandType =
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 0;
                connection.Open();
                var reader = cmd.ExecuteReader();

                var sheetNames = new List<string>();

                while (await reader.ReadAsync())
                {
                    sheetNames.Add(reader.GetValue(0).ToString());
                }

                while (await reader.NextResultAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        Console.WriteLine(reader.GetValue(0));
                    }
                }
                //do
                //{
                //    while (await reader.ReadAsync())
                //    {
                //        Console.WriteLine(reader.FieldCount);
                //        Console.WriteLine(reader["iduser"]);
                //        //do something with each record
                //    }
                //} while (await reader.NextResultAsync());
            }
        }

        private Dictionary<int, bool> SetQuotesIfText(SqlDataReader reader)
        {
            var columns = reader.FieldCount;

            var quotesNeeded = new Dictionary<int, bool>();

            for (int index = 0; index < columns; index++)
            {
                var typeName = reader.GetDataTypeName(index)
                    .ToLower();

                switch (typeName)
                {
                    case "varchar":
                    case "nvarchar":
                    case "char":
                    case "nchar":
                    case "text":
                    case "ntext":
                        quotesNeeded.Add(index, true);
                        break;
                    default:
                        quotesNeeded.Add(index, false);
                        break;
                }
            }

            return quotesNeeded;
        }

        private Dictionary<int, bool> SetToStringIfDate(SqlDataReader reader)
        {
            var columns = reader.FieldCount;

            var setToString = new Dictionary<int, bool>();

            for (int index = 0; index < columns; index++)
            {
                var typeName = reader.GetDataTypeName(index)
                    .ToLower();

                switch (typeName)
                {
                    case "datetime":
                    case "datetime2":
                    case "smalldatetime":
                    case "date":
                    case "timestamp":
                    case "time":
                        setToString.Add(index, true);
                        break;
                    default:
                        setToString.Add(index, false);
                        break;
                }
            }

            return setToString;
        }

        private double GetSumOfAddons(
            int idCarrier,
            int? idItalianCityFrom,
            int? idItalianCityTo,
            string zipFrom,
            string zipTo,
            int? idCountry,
            double? multipacketWeight,
            List<double?> multipacketDimensions,
            double initialPrice,
            List<addon_dimension> dbAddonsDimensions,
            List<addon_belt> dbAddonsBelts,
            List<addon_min_max_dim> dbAddonsMinMaxDim,
            List<addon_weight> dbAddonWeight,
            List<addon_sum_dim> dbAddonSum,
            List<addon_pricelist> dbAddonPricelist,
            List<addon_pricelist_zipcode> dbAddonPricelistZipCode,
            List<addon_city_zipcode> dbAddonCityZipCode,
            List<city> dbCities,
            List<addon_foreign_city_zipcode> dbAddonForeignCityZipCode,
            List<country> dbCountries,
            List<geonames_entity> dbGeonamesEntities)
        {
            var dimentionsAddon = GetDimensionAddon(idCarrier, multipacketDimensions, initialPrice, dbAddonsDimensions);
            var beltAddon = GetBeltPriceAddon(idCarrier, multipacketDimensions, initialPrice, dbAddonsBelts);
            var minMaxAddon = GetMinMaxDimAddon(idCarrier, multipacketDimensions, dbAddonsMinMaxDim);
            var weightAddon = GetWeightAddon(idCarrier, multipacketWeight, initialPrice, dbAddonWeight, true);
            var sumAddon = GetSumDimAddon(idCarrier, 0, multipacketDimensions[1], multipacketDimensions[2], multipacketDimensions[0], dbAddonSum);
            var cityZipCodeAddon = GetCityZipCodeAddOn(idCarrier, idItalianCityTo, idItalianCityFrom, "Pacchi",
                zipFrom, zipTo, dbAddonPricelist, dbAddonPricelistZipCode, dbAddonCityZipCode, dbCities);
            var foreignCityZipCodeAddon = GetForeignCityZipCodeAddOn(idCarrier, idItalianCityTo, idItalianCityFrom, "Pacchi",
                zipFrom, zipTo, idCountry, dbCities, dbAddonForeignCityZipCode, dbCountries, dbGeonamesEntities);

            var totalAddonPrice = 0.0;
            if (dimentionsAddon != null)
            {
                totalAddonPrice += dimentionsAddon.Price;
            }
            if (beltAddon != null)
            {
                totalAddonPrice += beltAddon.Price;
            }
            if (minMaxAddon != null)
            {
                totalAddonPrice += minMaxAddon.Price;
            }
            if (weightAddon != null)
            {
                totalAddonPrice += weightAddon.Price;
            }
            if (sumAddon != null)
            {
                totalAddonPrice += sumAddon.Price;
            }
            if (cityZipCodeAddon != null)
            {
                totalAddonPrice += cityZipCodeAddon.Price;
            }
            if (foreignCityZipCodeAddon != null)
            {
                totalAddonPrice += foreignCityZipCodeAddon.Price;
            }

            return totalAddonPrice;
        }

        private AddOnModel GetDimensionAddon(int idCarrier, List<double?> multipacketDimensions, double initialPrice, List<addon_dimension> dbAddonsDimensions)
        {
            var dimAddon = new AddOnModel();

            var orderedMultipacketDimensions = multipacketDimensions
                .OrderByDescending(d => d)
                .ToList();

            var max = orderedMultipacketDimensions[0];
            var med = orderedMultipacketDimensions[1];
            var min = orderedMultipacketDimensions[2];

            var dimensionAddOns =
                dbAddonsDimensions.Where
                (
                    x => x.carrierId == idCarrier && (max > x.max_dim || med > x.med_dim || min > x.min_dim)
                );

            if (dimensionAddOns.Any())
            {
                var dimensionAddOn =
                dimensionAddOns
                .OrderByDescending(x => x.max_dim)
                .ThenBy(x => x.med_dim)
                .ThenBy(x => x.min_dim)
                .FirstOrDefault();

                if (dimensionAddOn != null)
                {
                    if (dimensionAddOn.value != null)
                    {
                        dimAddon.Price += Convert.ToDouble(dimensionAddOn.value);
                        dimAddon.Cost += Convert.ToDouble(dimensionAddOn.internal_cost);
                    }
                    if (dimensionAddOn.percentage != null)
                    {
                        dimAddon.Price += Convert.ToDouble(dimensionAddOn.percentage) / 100 * Convert.ToDouble(initialPrice);
                        dimAddon.Cost += Convert.ToDouble(dimensionAddOn.percentage) / 100 * Convert.ToDouble(initialPrice);
                    }
                }
            }

            return dimAddon;
        }

        private AddOnModel GetBeltPriceAddon(int idCarrier, List<double?> multipacketDimensions, double initialPrice, List<addon_belt> dbAddonsBelts)
        {
            var orderedMultipacketDimensions = multipacketDimensions.OrderBy(x => x).ToList();

            var belt = (orderedMultipacketDimensions[0] + orderedMultipacketDimensions[1]) * 2 + orderedMultipacketDimensions[2];

            var getCarrierBelts = dbAddonsBelts.FirstOrDefault(x => x.carrierId == idCarrier && belt > x.minBelt && belt <= x.maxBelt);

            var beltAddon = new AddOnModel();
            if (getCarrierBelts != null)
            {
                if (getCarrierBelts.value != null)
                {
                    beltAddon.Price = Convert.ToDouble(getCarrierBelts.value);
                    beltAddon.Cost = Convert.ToDouble(getCarrierBelts.internal_cost);
                }
                if (getCarrierBelts.percentage != null)
                {
                    beltAddon.Price = initialPrice * Convert.ToDouble(getCarrierBelts.percentage) / 100;
                    beltAddon.Cost = initialPrice * Convert.ToDouble(getCarrierBelts.internal_cost) / 100;
                }
            }
            return beltAddon;
        }

        private AddOnModel GetMinMaxDimAddon(int idCarrier, List<double?> multipacketDimensions, List<addon_min_max_dim> dbAddonsMinMaxDim)
        {
            var orderedMultipacketDimensions = multipacketDimensions
             .OrderByDescending(d => d)
             .ToList();

            var max = orderedMultipacketDimensions[0];
            var med = orderedMultipacketDimensions[1];
            var min = orderedMultipacketDimensions[2];

            var addon = new AddOnModel();

            var getCarrierMinMaxDim = dbAddonsMinMaxDim.Where(x => x.idcarrier == idCarrier);
            if (getCarrierMinMaxDim.Any())
            {
                var minMaxSum = min + max;
                var minMaxData = getCarrierMinMaxDim.FirstOrDefault(x => x.min < minMaxSum && minMaxSum <= x.max);
                if (minMaxData != null)
                {
                    addon.Price = minMaxData.price;
                    addon.Cost = minMaxData.cost;
                }
            }

            return addon;
        }

        private AddOnModel GetWeightAddon(int idCarrier, double? weight, double initialPrice, List<addon_weight> dbAddonWeight, bool isSinglePackage = false)
        {
            var addon = new AddOnModel();

            if (!weight.HasValue)
            {
                return null;
            }

            var weightAddons = dbAddonWeight.Where(x =>
                       x.carrierid == idCarrier
                       && weight > x.weight
                       && (x.singlepkgonly == null || x.singlepkgonly.Value == isSinglePackage)
                        ).OrderByDescending(x => x.weight).FirstOrDefault();
            if (weightAddons != null)
            {
                if (weightAddons.value.HasValue)
                {
                    addon.Price += weightAddons.value.Value;
                    addon.Cost += weightAddons.internal_cost.Value;
                }
                else if (weightAddons.percentage.HasValue)
                {

                    addon.Price += weightAddons.percentage.Value / 100 * initialPrice;
                    addon.Cost += weightAddons.internal_cost.Value / 100 * initialPrice;
                }
                else if (weightAddons.everykg.HasValue)
                {
                    addon.Price += weightAddons.everykg.Value * (weight.Value - weightAddons.weight.Value);
                    addon.Cost += weightAddons.internal_cost.Value * (weight.Value - weightAddons.weight.Value);
                }
            }

            return addon;
        }

        public static AddOnModel GetSumDimAddon(int idcarrier, int iduser, double? h, double? w, double? d, List<addon_sum_dim> dbAddonSum)
        {
            var height = h.HasValue ? h.Value : 0;
            var width = w.HasValue ? w.Value : 0;
            var depth = d.HasValue ? d.Value : 0;

            var sumDim = height + width + depth;

            var addon = dbAddonSum.Where(
                a => a.idcarrier == idcarrier
                && sumDim >= a.sum_from
                && sumDim < a.sum_to
                && (height > a.min_side || width > a.min_side || depth > a.min_side))
                .FirstOrDefault();

            if (addon == null)
            {
                return null;
            }

            return new AddOnModel
            {
                Cost = addon.cost,
                Price = addon.price
            };
        }

        public static AddOnModel GetCityZipCodeAddOn(
            int idcarrier, int? idCityDestination,
            int? idCityDeparture, string serviceType,
            string capDeparture, string capDestination,
            List<addon_pricelist> dbAddonPricelist,
            List<addon_pricelist_zipcode> dbAddonPricelistZipCode,
            List<addon_city_zipcode> dbAddonCityZipCode,
            List<city> dbCities)
        {
            var addon = new AddOnModel();

            var addprfrom = (from x in dbAddonPricelist
                             where x.idcarrier == idcarrier
                             && (x.idcity == idCityDestination || x.idcity == idCityDeparture || x.idcity_from == idCityDeparture || x.idcity_to == idCityDestination)
                             group x by 1 into g
                             select new
                             {
                                 price = g.Sum(x => x.addon_price),
                                 internal_cost = g.Sum(x => x.addon_internal_cost)
                             }).FirstOrDefault();

            if (addprfrom != null)
            {
                addon.Price = addprfrom.price ?? 0;
                addon.Cost = addprfrom.internal_cost ?? 0;
            }

            if (idCityDestination == idCityDeparture)
            {
                addon.Price *= 2;
                addon.Cost *= 2;
            }

            if (serviceType.Equals("Pacchi") || serviceType.Equals("Lettere"))
            {
                int zipFrom, zipTo;
                if (!int.TryParse(capDeparture, out zipFrom))
                    return null;
                // throw new InvalidCastException("CAP mittente non presente o non valido");

                if (!int.TryParse(capDestination, out zipTo))
                    return null;
                //  throw new InvalidCastException("CAP destinatario non presente o non valido");

                var addOnByZipCode = dbAddonPricelistZipCode.Where(
                        x =>
                            x.idcarrier == idcarrier && (x.zipcode == zipFrom || x.zipcode == zipTo || x.zipcode_from == zipFrom || x.zipcode_to == zipTo));
                var addonZipPrice = Enumerable.Sum(addOnByZipCode, addonPricelistZipcode => addOnByZipCode != null ? Convert.ToDouble(addonPricelistZipcode.addon_price) : 0);
                var addonZipCost = Enumerable.Sum(addOnByZipCode, addonPricelistZipcode => addOnByZipCode != null ? Convert.ToDouble(addonPricelistZipcode.addon_internal_cost) : 0);

                if (zipFrom == zipTo)
                {
                    addonZipPrice *= 2;
                    addonZipCost *= 2;
                }

                addon.Price += addonZipPrice;
                addon.Cost += addonZipCost;

                if (dbAddonCityZipCode.Any(x => x.carrier_id == idcarrier))
                {
                    addon.Price = 0;
                    addon.Cost = 0;
                    var departureCity = dbCities.FirstOrDefault(x => x.idcity == idCityDeparture);
                    var destinationCity = dbCities.FirstOrDefault(x => x.idcity == idCityDestination);

                    addon_city_zipcode searchDepartureAddon = null;
                    if (departureCity != null)
                    {
                        searchDepartureAddon = dbAddonCityZipCode.FirstOrDefault(x =>
                        x.carrier_id == idcarrier && x.departure &&
                        (
                        (x.city.Equals(departureCity.name) && x.zipcode.Equals(capDeparture)) ||
                        (x.city.Equals(departureCity.name) && x.zipcode == null) ||
                        (x.city == null && x.zipcode.Equals(capDeparture))
                        ));
                    }

                    addon_city_zipcode searchDestinationAddon = null;
                    if (destinationCity != null)
                    {
                        searchDestinationAddon = dbAddonCityZipCode.FirstOrDefault(x =>
                        x.carrier_id == idcarrier && x.destination &&
                        (
                            (x.city != null && x.city.Equals(destinationCity.name) && (x.zipcode != null && x.zipcode.Equals(capDestination))) ||
                            (x.city != null && x.city.Equals(destinationCity.name) && x.zipcode == null) ||
                            (x.city == null && x.zipcode.Equals(capDestination))
                        ));
                    }

                    addon.Price = searchDepartureAddon != null ? searchDepartureAddon.addon_price.Value : 0;
                    addon.Price += searchDestinationAddon != null ? searchDestinationAddon.addon_price.Value : 0;

                    addon.Cost = searchDepartureAddon != null ? searchDepartureAddon.addon_internal_cost.Value : 0;
                    addon.Cost += searchDestinationAddon != null ? searchDestinationAddon.addon_internal_cost.Value : 0;
                }
            }

            return addon;
        }

        public static AddOnModel GetForeignCityZipCodeAddOn(int idcarrier, int? idCityDestination, int? idCityDeparture,
            string serviceType, string capDeparture, string capDestination, int? countryid, List<city> dbCities,
            List<addon_foreign_city_zipcode> dbAddonForeignCityZipCode,
            List<country> dbCountries, List<geonames_entity> dbGeonamesEntity)
        {
            var addon = new AddOnModel();
            if (!(serviceType.Equals("Pacchi") || serviceType.Equals("Lettere")))
                return addon;

            using (var sb = new SendaboxEntities())
            {
                if (dbAddonForeignCityZipCode.Any(x => x.carrier_id == idcarrier))
                {
                    addon.Price = 0;
                    addon.Cost = 0;
                    var departureCity = dbCities.FirstOrDefault(x => x.idcity == idCityDeparture);

                    /* DEPARTURE */
                    addon_foreign_city_zipcode searchDepartureAddon = null;
                    var searchDepartureAddonQuery = dbAddonForeignCityZipCode.Where(x =>
                        x.carrier_id == idcarrier && x.idCountry == 1 && x.departure);
                    if (!string.IsNullOrEmpty(capDeparture))
                    {
                        searchDepartureAddon = searchDepartureAddonQuery.FirstOrDefault(
                                x => ((capDeparture == x.zipcode) && (x.city == null || x.city == departureCity.name))
                        );
                    };
                    if (searchDepartureAddon == null)
                    {
                        searchDepartureAddon = searchDepartureAddonQuery.FirstOrDefault(x =>
                            (x.zipcode == null) && (x.city == departureCity.name));
                    }

                    /* DESTINATION */
                    //  -- Check Country
                    country country_to = null;

                    country_to = dbCountries.First(x => x.idcountry == countryid);

                    if (country_to == null)
                        throw new Exception("Impossibile to look up country");
                    string cityName = null;

                    if (country_to.IsSourceGeoNames() && idCityDestination.HasValue)
                    {
                        //GR 16/02/2021 VIA API DA ERRORE
                        var gne = dbGeonamesEntity.FirstOrDefault(x => x.geonameId == idCityDestination);
                        if (gne != null)
                        {
                            cityName = gne.asciiName;
                        }

                    }
                    else if (idCityDestination.HasValue)
                    {
                        cityName = dbCities.FirstOrDefault(x => x.idcity == idCityDestination).name;
                    }

                    if (cityName == null)
                    {
                        return null;
                    }

                    addon_foreign_city_zipcode searchDestinationAddon = null;
                    var searchDestinationAddonQuery = dbAddonForeignCityZipCode.
                        Where(x => x.carrier_id == idcarrier && x.idCountry == countryid && x.destination);
                    if (!string.IsNullOrEmpty(capDestination))
                    {
                        searchDestinationAddon = searchDestinationAddonQuery.FirstOrDefault(
                            x => ((x.zipcode.CompareTo(capDestination) <= 0 && capDestination.CompareTo(x.zipcode_to) <= 0) && (x.city == null || x.city == cityName))
                                );
                    };
                    if (searchDestinationAddon == null)
                    {
                        searchDestinationAddon = searchDestinationAddonQuery.FirstOrDefault(x =>
                           (x.zipcode == null) && (x.city == cityName));
                    }

                    addon.Price = searchDepartureAddon != null ? searchDepartureAddon.addon_price.Value : 0;
                    addon.Price += searchDestinationAddon != null ? searchDestinationAddon.addon_price.Value : 0;

                    addon.Cost = searchDepartureAddon != null ? searchDepartureAddon.addon_internal_cost.Value : 0;
                    addon.Cost += searchDestinationAddon != null ? searchDestinationAddon.addon_internal_cost.Value : 0;
                }

            }
            return addon;
        }


        internal DailyExportFileDto MapDailyExportDbToFileDto(DailyExportDbDto dbExport, double addon)
        {
            return new DailyExportFileDto
            {
                IdUserCityOrder = dbExport.IdUserCityOrder,
                IdUser = dbExport.IdUser,
                Login = dbExport.Login,
                OrderPrice = dbExport.OrderPrice,
                Sales = dbExport.Sales,
                OrderDate = dbExport.OrderDate,
                FromProvince = dbExport.FromProvince,
                ToCountry = dbExport.ToCountry,
                ToRegion = dbExport.ToRegion,
                Weight = dbExport.Weight,
                WeightRange = dbExport.WeightRange,
                WeightUpperRange = dbExport.WeightUpperRange,
                Agent = dbExport.Agent,
                BuisnessName = dbExport.BuisnessName,
                Status = dbExport.Status,
                TargetCountry = dbExport.ToCountry,
                OrderMonth = dbExport.OrderMonth,
                OrderWeek = dbExport.OrderWeek,
                PaymentMethod = dbExport.PaymentMethod,
                EstimatedCost = dbExport.EstimatedCost,
                Carrier = dbExport.Carrier,
                CustomerTracking = dbExport.CustomerTracking,
                PrivateCompany = dbExport.PrivateCompany,
                CouponName = dbExport.CouponName,
                Addon = addon
            };
        }

        private List<addon_dimension> AllDbAddonDimensions()
        {
            using (var db = new SendaboxEntities())
            {
                return db.addon_dimension
                 .AsNoTracking()
                 .ToList();
            }
        }

        private List<addon_belt> AllDbAddonBelts()
        {
            using (var db = new SendaboxEntities())
            {
                return db.addon_belt
                 .AsNoTracking()
                 .ToList();
            }
        }

        private List<addon_min_max_dim> AllDbAddonMinMaxDim()
        {
            using (var db = new SendaboxEntities())
            {
                return db.addon_min_max_dim
                 .AsNoTracking()
                 .ToList();
            }
        }

        private List<addon_weight> AllDbAddonWeight()
        {
            var addon = new AddOnModel();
            using (var sb = new SendaboxEntities())
            {
                return sb.addon_weight
                    .AsNoTracking()
                    .ToList();
            }
        }

        private List<addon_sum_dim> AllDbAddonSum()
        {
            var addon = new AddOnModel();
            using (var sb = new SendaboxEntities())
            {
                return sb.addon_sum_dim
                    .AsNoTracking()
                    .ToList();
            }
        }

        private List<addon_pricelist> AllDbAddonPricelist()
        {

            using (var sb = new SendaboxEntities())
            {
                return sb.addon_pricelist
                    .AsNoTracking()
                    .ToList();
            }
        }

        private List<addon_pricelist_zipcode> AllDbAddonPricelistZipCode()
        {
            var addon = new AddOnModel();
            using (var sb = new SendaboxEntities())
            {
                return sb.addon_pricelist_zipcode
                    .AsNoTracking()
                    .ToList();
            }
        }

        private List<addon_city_zipcode> AllDbAddonCityZipCode()
        {
            var addon = new AddOnModel();
            using (var sb = new SendaboxEntities())
            {
                return sb.addon_city_zipcode
                    .AsNoTracking()
                    .ToList();
            }
        }

        private List<city> AllDbCities()
        {
            var addon = new AddOnModel();
            using (var sb = new SendaboxEntities())
            {
                return sb.cities
                    .AsNoTracking()
                    .ToList();
            }
        }

        private List<addon_foreign_city_zipcode> AllDbForeignCityZipCode()
        {
            var addon = new AddOnModel();
            using (var sb = new SendaboxEntities())
            {
                return sb.addon_foreign_city_zipcode
                    .AsNoTracking()
                    .ToList();
            }
        }

        private List<country> AllDbCountries()
        {
            var addon = new AddOnModel();
            using (var sb = new SendaboxEntities())
            {
                return sb.countries
                    .AsNoTracking()
                    .ToList();
            }
        }

        private List<geonames_entity> AllDbGeonamesEntity()
        {
            var addon = new AddOnModel();
            using (var sb = new SendaboxEntities())
            {
                return sb.geonames_entity
                    .AsNoTracking()
                    .ToList();
            }
        }
    }
}