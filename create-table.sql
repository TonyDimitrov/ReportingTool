
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[EXTRACTION_REPORT](
	[id_extraction_data] [int] IDENTITY(1,1) NOT NULL,
	[extraction_storedprocedure] [varchar](50) NULL,
	[extraction_email] [varchar](500) NULL,
	[email_subject] [varchar](100) NULL,
	[extraction_format] [varchar](10) NULL,
	[file_name] [varchar](100) NULL,
	[email_text] [varchar](1000) NULL,
	[is_active] [bit] NULL,
	[send_compressed] [bit] NULL,
	[send_empty] [bit] NULL,
	[send_on_business_days] [bit] NULL,
	[send_on_week_days] [varchar](200) NULL,
	[send_on_month_beginning] [bit] NULL,
	[send_on_month_end] [bit] NULL,
	[ignore_from_date] [date] NULL,
	[ignore_to_date] [date] NULL,
	[send_by_protocol] [smallint] NULL,
	[ftp_host] [varchar](100) NULL,
	[ftp_port] [smallint] NULL,
	[ftp_username] [varchar](50) NULL,
	[ftp_password] [varchar](50) NULL,
	[ftp_remote_folder] [varchar](200) NULL,
	[type_of_excel] [smallint] NULL,
 CONSTRAINT [PK_EXTRACTION_REPORT] PRIMARY KEY CLUSTERED 
(
	[id_extraction_data] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


