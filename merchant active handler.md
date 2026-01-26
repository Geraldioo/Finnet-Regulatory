nah sekarang kita sudah simpen ke database saya ingin mempuat flow baru dari no 1 yaitu untuk mengisi template master data merchant active

sumber datanya


SELECT *
FROM master_data
WHERE master_data_type = 'autoreport_merchant_active_qris';
# ID	transaction_date	merchant_city	keterangan	merchant_category	merchant_criteria	jumlah_trx	amount_trx	merchant_category_name	nama_kab_kota	master_merchant_city	created_at	master_data_type
10621	2026-01-01	ACEH SINGKIL		5812	UKE		1	Regular	ACEH SINGKIL		2026-01-11 12:54:22	autoreport_merchant_active_qris
10622	2026-01-01	AMBON		5331	UKE		1	Regular	AMBON	Kota Ambon	2026-01-11 12:54:22	autoreport_merchant_active_qris
10623	2026-01-01	ASAHAN		5812	UMI		1	Regular	ASAHAN	kab. asahan	2026-01-11 12:54:22	autoreport_merchant_active_qris
10624	2026-01-01	ASAHAN		5945	UMI		1	Regular	ASAHAN	kab. asahan	2026-01-11 12:54:22	autoreport_merchant_active_qris
10625	2026-01-01	ASAHAN		8661	URE		1	Donasi Sosial	ASAHAN	kab. asahan	2026-01-11 12:54:22	autoreport_merchant_active_qris
10626	2026-01-01	BADUNG		3333	UKE		1	Regular	BADUNG	Kab. Badung	2026-01-11 12:54:22	autoreport_merchant_active_qris
10627	2026-01-01	BADUNG		5331	UKE		2	Regular	BADUNG	Kab. Badung	2026-01-11 12:54:22	autoreport_merchant_active_qris
10628	2026-01-01	BADUNG		5812	UKE		6	Regular	BADUNG	Kab. Badung	2026-01-11 12:54:22	autoreport_merchant_active_qris


ini direcotry excelnya
C:\RPA\FINNETDEV\Usecase 1- Qris Report\Qris-Report\Master Data\3. merchant_aktif_qris2025.xslx

diisi bagian sheet merchant_qris (version 1)

ini isi excelnya/header dan data
Date	Kota	MCC	Kategori	Kriteria	Jumlah	Kota	Kota fix


1. Configure MySQL Database
   Database address: "localhost"
   port: 3306
   username: "root"
   password: 12345
   database name: "regulatorydev"
   Database configuration object: regulatorydbdev
   
   Console Log: "Starting Master Data Merchant Active Excel population process"

2. Assign Value To Variable (Excel Template Path)
   Value: "C:\\RPA\\FINNETDEV\\Usecase 1- Qris Report\\Qris-Report\\Master Data\\3. merchant_aktif_qris2025.xlsx"
   Variable name: masterDataExcelPath

3. Assign Value To Variable (Target Sheet Name)
   Value: "merchant_qris (version 1)"
   Variable name: targetSheetName

4. Get Current Date and Time
   current datetime: currentDateTime

5. Date To String
   Datetime Object: currentDateTime
   Output string format: YYYY-MM-DD
   Conversion Result: todayDate

6. Console Log: "Fetching data from database for date: " + todayDate

7. Execute SQL Statements (Call Stored Procedure)
   SQL Statements: ["CALL GetMerchantActiveForExcel('" + todayDate + "')"]
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


CREATE DEFINER=`root`@`localhost` PROCEDURE `GetMerchantActiveForExcel`(
    IN p_report_date DATE
)
BEGIN
    -- Hitung Kamis minggu lalu dan Rabu kemarin
    DECLARE v_start_date DATE;
    DECLARE v_end_date DATE;
    
    SET v_end_date = DATE_SUB(p_report_date, INTERVAL 1 DAY);
    SET v_start_date = DATE_SUB(p_report_date, INTERVAL 7 DAY);
    
    SELECT 
        DATE_FORMAT(transaction_date, '%Y-%m-%d') AS 'Date',
        merchant_city AS 'Kota',
        merchant_category AS 'MCC',
        COALESCE(merchant_category_name, '') AS 'Kategori',
        merchant_criteria AS 'Kriteria',
        amount_trx AS 'Jumlah',
        COALESCE(nama_kab_kota, '') AS 'Kota 2',
        COALESCE(master_merchant_city, '') AS 'Kota fix'
    FROM master_data
    WHERE master_data_type = 'autoreport_merchant_active_qris'
      AND DATE(created_at) >= v_start_date
      AND DATE(created_at) <= v_end_date
    ORDER BY transaction_date, merchant_city;
    
END