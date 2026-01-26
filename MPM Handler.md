nah sekarang kita sudah simpen ke database saya ingin mempuat flow baru dari no 1 yaitu untuk mengisi template master data MPM

sumber datanya

SELECT *
FROM master_data
WHERE master_data_type = 'autoreport_qris_acquirer_off_us_mpm';
# ID	transaction_date	merchant_city	keterangan	merchant_category	merchant_criteria	jumlah_trx	amount_trx	merchant_category_name	nama_kab_kota	master_merchant_city	created_at	master_data_type
1	2026-01-06	ACEH SINGKIL 	ACQUIRER_OFFUS	5812	UKE	3	54000	Regular	ACEH SINGKIL 		2026-01-07 15:32:42	autoreport_qris_acquirer_off_us_mpm
2	2026-01-06	ASAHAN       	ACQUIRER_OFFUS	5621	UMI	1	790000	Regular	ASAHAN       	kab. asahan	2026-01-07 15:32:42	autoreport_qris_acquirer_off_us_mpm
3	2026-01-06	ASAHAN       	ACQUIRER_OFFUS	5945	UMI	1	100000	Regular	ASAHAN       	kab. asahan	2026-01-07 15:32:42	autoreport_qris_acquirer_off_us_mpm

SELECT *
FROM master_data
WHERE master_data_type = 'autoreport_qris_issuer_off_us_mpm';
# ID	transaction_date	merchant_city	keterangan	merchant_category	merchant_criteria	jumlah_trx	amount_trx	merchant_category_name	nama_kab_kota	master_merchant_city	created_at	master_data_type
'1113', '2026-01-06', 'JAKARTA SELAT', 'ISSUER OFFUS', '5816', 'UBE', '2', '1099850', NULL, 'JAKARTA SELAT', 'Wil. Kota Jakarta Selatan', '2026-01-07 15:33:18', 'autoreport_qris_issuer_off_us_mpm'
'1114', '2026-01-06', 'PURWAKARTA KA', 'ISSUER OFFUS', '5462', 'UME', '1', '22000', 'Regular', 'PURWAKARTA KA', 'kab. purwakarta', '2026-01-07 15:33:18', 'autoreport_qris_issuer_off_us_mpm'
'1115', '2026-01-06', 'MEDAN', 'ISSUER OFFUS', '5411', 'UKE', '1', '100000', 'Regular', 'MEDAN', 'Kota Medan', '2026-01-07 15:33:18', 'autoreport_qris_issuer_off_us_mpm'
'1116', '2026-01-06', 'PURWAKARTA', 'ISSUER OFFUS', '5814', 'UBE', '1', '15000', 'Regular', 'PURWAKARTA', 'Kab. Purwakarta', '2026-01-07 15:33:18', 'autoreport_qris_issuer_off_us_mpm'

SELECT *
FROM master_data
WHERE master_data_type = 'autoreport_qris_issuer_on_us_mpm';
# ID	transaction_date	merchant_city	keterangan	merchant_category	merchant_criteria	jumlah_trx	amount_trx	merchant_category_name	nama_kab_kota	master_merchant_city	created_at	master_data_type
1263	2026-01-06	ACEH SINGKIL	ISSUER_ONUS	5812	UKE	4	85000	Regular	ACEH SINGKIL		2026-01-07 15:34:06	autoreport_qris_issuer_on_us_mpm
1264	2026-01-06	ASAHAN	ISSUER_ONUS	5621	UMI	1	790000	Regular	ASAHAN	kab. asahan	2026-01-07 15:34:06	autoreport_qris_issuer_on_us_mpm
1265	2026-01-06	ASAHAN	ISSUER_ONUS	5945	UMI	1	100000	Regular	ASAHAN	kab. asahan	2026-01-07 15:34:06	autoreport_qris_issuer_on_us_mpm



ini direcotry excelnya
C:\RPA\FINNETDEV\Usecase 1- Qris Report\Qris-Report\Master Data\1. Master Data Pelaporan Transaksi QRIS MPM  2025.xslx

diisi bagian sheet New Query

ini isi excelnya/header
transaction_date,	merchant_city,	Fungsi PJSP,	merchant_category,	merchant_criteria,	jml_trx,	amount_trx,	Merchant Category Name,	Nama Kota/Kab, Merchant City 2






1. Configure MySQL Database
   Database address: "localhost"
   port: 3306
   username: "root"
   password: 12345
   database name: "regulatorydev"
   Database configuration object: regulatorydbdev
   
   Console Log: "Starting Master Data MPM Excel population process"

2. Assign Value To Variable (Excel Template Path)
   Value: "C:\\RPA\\FINNETDEV\\Usecase 1- Qris Report\\Qris-Report\\Master Data\\1. Master Data Pelaporan Transaksi QRIS MPM 2025.xlsx"
   Variable name: masterDataExcelPath

3. Assign Value To Variable (Target Sheet Name)
   Value: "New Query"
   Variable name: targetSheetName

4. Get Current Date and Time
   current datetime: currentDateTime

5. Date To String
   Datetime Object: currentDateTime
   Output string format: YYYY-MM-DD
   Conversion Result: todayDate

6. Console Log: "Fetching data from database for date: " + todayDate

7. Execute SQL Statements (Call Stored Procedure)
   SQL Statements: ["CALL GetMasterDataMPMForExcel('" + todayDate + "')"]
   Database config: regulatorydbdev
   output format: Sqltable
   timeout (second): 300
   Query result: masterDataResult (array<cyclone.Sqltable>)

8. Console Log: "Data fetched successfully"

9. Get Array Length
   Array: masterDataResult
   Array length: totalResultSets

10. Console Log: "Total result sets: " + totalResultSets

11. If - Conditional judgment:
    IF totalResultSets > 0
    THEN:
       11.1 Get Array Elements
            Array: masterDataResult
            Target Element Subscript: 0
            Result: firstSqlTable (cyclone.Sqltable)

       11.2 Convert to Array
            Data Table: firstSqlTable
            Include Table Header: No
            Conversion result: excelDataArray (array<array>)

       11.3 Get Array Length
            Array: excelDataArray
            Array length: totalRows

       11.4 Console Log: "Data transformation completed. Rows: " + totalRows

       11.5 If - Conditional judgment:
            IF totalRows > 0
            THEN:
               11.5.1 Open Excel Workbook
                      File path: masterDataExcelPath
                      When the file does not exist: Do not automatically create
                      Open mode: automatic detection
                      Is it visible: No
                      Excel File: masterDataExcelObj

               11.5.2 Console Log: "Excel file opened: " + masterDataExcelPath

               11.5.3 Get the Number of Rows and Columns
                      Excel File Object: masterDataExcelObj
                      Worksheet: targetSheetName
                      Get Rows/Columns: Rows
                      Column Number: "A" (atau 1)
                      Number of Rows: usedRowCount

               11.5.4 Console Log: "Current used rows: " + usedRowCount

               11.5.5 Assign Value To Variable (Calculate Next Row)
                      Value: usedRowCount + 1
                      Variable name: nextRow

               11.5.6 Assign Value To Variable (Build Starting Cell)
                      Value: "A" + nextRow
                      Variable name: startingCell

               11.5.7 Console Log: "Writing data starting from: " + startingCell

               11.5.8 Write Range
                      Excel file object: masterDataExcelObj
                      Select worksheet name/serial number: targetSheetName
                      Starting cell: By cell name
                      Cell name: startingCell
                      Data format: General
                      Variable type: Two-dimensional array
                      Range data: excelDataArray
                      Ignore header/first row: Not to remove
                      Whether to automatically save: No

               11.5.9 Console Log: "Data written to sheet: " + targetSheetName + " (" + totalRows + " rows)"

               11.5.10 Close Excel Workbook
                       Excel file object: masterDataExcelObj
                       Closing Method: Save to Original Path and Close

               11.5.11 Console Log: "Excel file saved and closed"

            ELSE:
               Console Log: "No data rows to write"

    ELSE:
       Console Log: "No data found for date: " + todayDate

12. Console Log: "Completed Successfully"


CREATE DEFINER=`root`@`localhost` PROCEDURE `GetMasterDataMPMForExcel`(
    IN p_report_date DATE
)
BEGIN
    -- Hitung Kamis minggu lalu dan Rabu kemarin
    -- Contoh: Jika p_report_date = 2026-01-15 (Kamis)
    --         Start = 2026-01-08 (Kamis minggu lalu)
    --         End   = 2026-01-14 (Rabu kemarin)
    
    DECLARE v_start_date DATE;
    DECLARE v_end_date DATE;
    
    -- End date = Rabu kemarin (1 hari sebelum Kamis hari ini)
    SET v_end_date = DATE_SUB(p_report_date, INTERVAL 1 DAY);
    
    -- Start date = Kamis minggu lalu (7 hari sebelum Kamis ini)
    SET v_start_date = DATE_SUB(p_report_date, INTERVAL 7 DAY);
    
    SELECT 
        DATE_FORMAT(transaction_date, '%Y-%m-%d') AS transaction_date,
        merchant_city,
        keterangan AS 'Fungsi PJSP',
        merchant_category,
        merchant_criteria,
        jumlah_trx AS jml_trx,
        amount_trx,
        COALESCE(merchant_category_name, '') AS 'Merchant Category Name',
        COALESCE(nama_kab_kota, '') AS 'Nama Kota/Kab.',
        COALESCE(master_merchant_city, '') AS 'Merchant City 2'  -- ⭐ TAMBAHAN
    FROM master_data
	WHERE master_data_type IN (
	'autoreport_qris_acquirer_off_us_mpm',
	'autoreport_qris_issuer_off_us_mpm',
	'autoreport_qris_issuer_on_us_mpm'
    )
      AND DATE(created_at) >= v_start_date
      AND DATE(created_at) <= v_end_date
    ORDER BY transaction_date, merchant_city;
    
END








1. Prepare and open Master_data Excel File
2. Hit Stored procedure untuk populate data master nyadata master yang didapat di olah dan di buat menjadi nested array yang struktur kolom nya mengikuti file excel
3. Paste Nested array data ke file master_data terkait