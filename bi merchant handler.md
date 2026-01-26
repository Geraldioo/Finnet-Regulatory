nah sekarang kita sudah simpen ke database saya ingin mempuat flow baru dari no 1 yaitu untuk mengisi template master data bi merchant / merchant cpm

sumber datanya

SELECT * FROM master_data_merchant_cpm;
# id	nama_merchant_50	nama_merchant_25	merchant_id	merchant_criteria	merchant_category	status	periode_aktivasi	merchant_category_name	created_at
1	HENGKI	HENGKI	MRC2024063008381367911	UKE	7538	Active	2024-07-12 10:50:29	Regular	2026-01-08 20:44:32
2	Marista Nurlina Hutauruk	Marista Nurlina Hutauruk	MRC2025110417483725891	BLU	8099	Active	2025-11-14 09:43:13	Regular	2026-01-11 13:51:55
3	Syukron Maulana	Syukron Maulana	MRC2024052604302768978	UKE	5945	Active	2024-06-01 08:57:42	Regular	2026-01-13 10:05:49
4	BOY SE	BOY SE	MRC2024071402340857691	UKE	5978	Active	2024-10-23 08:28:42	Regular	2026-01-13 10:05:49
5	Nanang misbachul munir	Nanang misbachul munir	MRC2024080404590319296	UKE	5965	Active	2024-08-06 09:33:51	Regular	2026-01-13 10:05:49
6	anita trisnawati	anita trisnawati	MRC2025091611494136429	UKE	5499	Active	2025-09-29 14:38:45	Regular	2026-01-14 15:40:23


ini direcotry excelnya
C:\RPA\FINNETDEV\Usecase 1- Qris Report\Qris-Report\Master Data\4. Master Merchant CPM.xslx

diisi bagian sheet List Merchant

ini isi excelnya/header dan data
| Row | A (No) | B (NMID) | C Nama Merchant (Max 50)        | D Nama Merchant (Max 25)       | E MPA | F MID                   | G Kota            | H Kode | I Kriteria | J MCC | K Jml Term | L Jml Mer | M NPW | N KTP     | O Tipe C | P Status | Q Periode Aktivasi | R Category Merchant |
|-----|--------|----------|--------------------------------|--------------------------------|-------|--------------------------|-------------------|--------|------------|-------|------------|------------|--------|-----------|----------|----------|--------------------|-------------------|
| 5   | 1      |          | Alexander simamora             | Alexander simamora             |       | MRC2023030210451777859   | BEKASI            | 17510  | UKE        | 5977  | 1          | 1          |        | 64710108  | B        | Active   | 2 - 8 Maret 2023 | Regular           |
| 6   |        |          | ade santosa                    | ade santosa                    |       | MRC2023030214523490373   | TANGERANG         | 15145  | UME        | 7531  |            |            |        |           |          | Active   | 2 - 8 Maret 2023 | Regular           |
| 7   |        |          | BENY SISWANTO                  | BENY SISWANTO                  |       | MRC2023030508014550067   | INDRAMAYU         | 45273  | UMI        | 5641  |            |            |        |           |          | Active   | 2 - 8 Maret 2023 | Regular           |



1. Configure MySQL Database
   Database address: "localhost"
   port: 3306
   username: "root"
   password: 12345
   database name: "regulatorydev"
   Database configuration object: regulatorydbdev
   
   Console Log: "Starting Master Data Merchant MPM population process"

2. Assign Value To Variable (Excel Template Path)
   Value: "C:\\RPA\\FINNETDEV\\Usecase 1- Qris Report\\Qris-Report\\Master Data\\4. Master Merchant CPM.xlsx"
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
   SQL Statements: ["CALL GetMerchantCPMForExcel('" + todayDate + "')"]
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


CREATE DEFINER=`root`@`localhost` PROCEDURE `GetMerchantCPMForExcel`(
    IN p_report_date DATE
)
BEGIN
    DECLARE v_start_date DATE;
    DECLARE v_end_date DATE;
    DECLARE periode_aktivasi_text VARCHAR(100);
   
    -- Hitung range 7 hari sebelum p_report_date
    SET v_end_date   = DATE_SUB(p_report_date, INTERVAL 1 DAY);
    SET v_start_date = DATE_SUB(p_report_date, INTERVAL 7 DAY);
   
    -- Buat string periode dalam format Indonesia: "%e - %e %M %Y"
    -- Contoh: "8 - 14 Januari 2025"
    SET periode_aktivasi_text = CONCAT(
        DATE_FORMAT(v_start_date, '%e'), ' - ',
        DATE_FORMAT(v_end_date,   '%e %M %Y')
    );
   
    -- Row number
    SET @row_number = 0;
   
    SELECT
        (@row_number := @row_number + 1) AS 'No',
        '' AS 'NMID',
        nama_merchant_50 AS 'Nama Merchant (Max 50)',
        nama_merchant_25 AS 'Nama Merchant (Max 25)',
        '' AS 'MPA',
        merchant_id AS 'MID',
        city AS 'Kota',
        kode_pos AS 'Kode',
        merchant_criteria AS 'Kriteria',
        merchant_category AS 'MCC',
        '' AS 'Jml Term',
        '' AS 'Jml Mer',
        '' AS 'NPW',
        '' AS 'KTP',
        '' AS 'Tipe C',
        status AS 'Status',
        periode_aktivasi_text AS 'Periode Aktivasi',   -- ← GANTI JADI STRING STATIS INI
        COALESCE(merchant_category_name, '') AS 'Category Merchant'
    FROM master_data_merchant_cpm
    WHERE DATE(created_at) >= v_start_date
      AND DATE(created_at) <= v_end_date
    ORDER BY created_at, merchant_id;
   
END