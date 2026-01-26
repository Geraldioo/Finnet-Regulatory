regulatory Qris 2nd


kita mendownload attachments dari email berupa file excel ada 8 email dan 8 attachments yg berbeda2 isinya
1. autoreport_qris_acquirer_off_us_mpm
isi datannya:
tanggal	merchant_city	ket	merchant	merchant_criteria	jml_trx	amount_trx
2026-01-07	ACEH SINGKIL 	ACQUIRER_OFFUS	5812	UKE	1	11000
2026-01-07	AMBON        	ACQUIRER_OFFUS	7941	UKE	1	200000
2026-01-07	ASAHAN       	ACQUIRER_OFFUS	5945	UMI	1	110000
...
2. autoreport_qris_issuer_off_us_mpm
isi datannya:
tanggal	merchant_city	ket	merchant	merchant_criteria	jml_trx	amount_trx
07-01-2026	JAKARTA PUSAT	ISSUER OFFUS	4816	UKE	1	50000
07-01-2026	PEKANBARU	ISSUER OFFUS	4816	UME	1	50000
07-01-2026	JAKARTA SELAT	ISSUER OFFUS	4722	UME	1	25643
...
3. autoreport_qris_issuer_on_us_mpm
isi datannya:
tanggal	merchant_city	ket	merchant	merchant_criteria	jml_trx	amount_trx
2026-01-07	ACEH SINGKIL	ISSUER_ONUS	5812	UKE	1	11000
2026-01-07	AMBON	ISSUER_ONUS	7941	UKE	3	225000
2026-01-07	ASAHAN	ISSUER_ONUS	5051	URE	4	5551740
...
4. autoreport_qris_acquirer_off_us_cpm
isi datannya:
tanggal	merchant_city	ket	merchant	merchant_criteria	jml_trx	amount_trx
01-01-2026	PANDEGLANG	ACQUIRER_OFFUS	6013	UKE	117	424500001
01-01-2026	DEPOK	ACQUIRER_OFFUS	6013	UKE	54	128300000
01-01-2026	SUMEDANG	ACQUIRER_OFFUS	6013	UKE	3	10500002
...
5. autoreport_qris_issuer_off_us_cpm
isi datannya:
tanggal	merchant_city	ket	merchant	merchant_criteria	jml_trx	amount_trx
01-01-2026	DEPOK	ISSUER OFFUS	6013	UKE	5	16310000
01-01-2026	BOGOR	ISSUER OFFUS	6013	UKE	9	17400000
01-01-2026	JAKARTA UTARA	ISSUER OFFUS	5411	UBE	2	319300
...
6. autoreport_qris_issuer_on_us_cpm
isi datannya:
DATE_FORMAT(trx.trx_dtime, '%d-%m-%Y')	merchant_city	ket	merchant	merchant_criteria	jml_trx	amount_trx
01-01-2026	DEPOK	ISSUER_ONUS	5499	UKE	5	16310000
01-01-2026	BOGOR	ISSUER_ONUS	5499	UKE	9	17400000
01-01-2026	PANDEGLANG	ISSUER_ONUS	5499	UKE	12	19800005
...
7. autoreport_bi_merchant_qris
isi datannya:
merchant_id	pic_name	mcc	merchant_criteria	province	city	district	village	kodepos	approval_date
MRC2024063008381367911	HENGKI	7538	UKE	SUMATERA SELATAN	OGAN KOMERING ILIR	TULUNG SELAPAN	TULUNG SELAPAN ILIR	30655	2024-07-12 10:50:29
...
8. autoreport_merchant_active_qris
isi datannya:
tanggal	city	merchant	merchant_criteria	total
2026-01-01	ACEH SINGKIL	5812	UKE	1
2026-01-01	AMBON	5331	UKE	1
2026-01-01	ASAHAN	5812	UMI	1
...

ini skema databasenya:
tabel master_data
'ID', 'int', 'NO', 'PRI', NULL, 'auto_increment'
'transaction_date', 'date', 'YES', '', NULL, ''
'merchant_city', 'varchar(50)', 'YES', '', NULL, ''
'keterangan', 'varchar(50)', 'YES', '', NULL, ''
'merchant_category', 'int', 'YES', '', NULL, ''
'merchant_criteria', 'varchar(50)', 'YES', '', NULL, ''
'jumlah_trx', 'int', 'YES', '', NULL, ''
'amount_trx', 'bigint', 'YES', '', NULL, ''
'merchant_category_name', 'varchar(50)', 'YES', '', NULL, ''
'nama_kab_kota', 'varchar(50)', 'YES', '', NULL, ''
'master_merchant_city', 'varchar(50)', 'YES', '', NULL, ''
'created_at', 'datetime', 'YES', '', NULL, ''
'master_data_type', 'varchar(50)', 'YES', '', NULL, ''


tabel master_data_merchant_cpm
# Field	Type	Null	Key	Default	Extra
id	int	NO	PRI		auto_increment
nama_merchant_50	varchar(50)	YES			
nama_merchant_25	varchar(25)	YES			
merchant_id	varchar(50)	YES			
merchant_criteria	varchar(50)	YES			
merchant_category	int	YES			
city	varchar(50)	YES			
kode_pos	varchar(10)	YES			
status	varchar(50)	YES			
periode_aktivasi	varchar(50)	YES			
merchant_category_name	varchar(50)	YES			
created_at	datetime	YES			


store procedure nya
CREATE DEFINER=`root`@`localhost` PROCEDURE `InsertQrisRegulatoryData`(
    IN p_dataType VARCHAR(50),
    IN p_content LONGTEXT
)
BEGIN
    DECLARE i INT DEFAULT 1;
    DECLARE row_count INT DEFAULT 0;
    DECLARE start_pos INT DEFAULT 3;
    DECLARE end_pos INT;
    DECLARE current_row TEXT;
    DECLARE col_count INT DEFAULT 0;
    
    -- Variables untuk MPM data (7 kolom)
    DECLARE date_str VARCHAR(20);
    DECLARE v_transaction_date DATE;
    DECLARE col2, col3, col4, col5, col6, col7 VARCHAR(255);
    DECLARE v_jumlah_trx INT;
    DECLARE v_amount_trx DECIMAL(18,0);
    
    -- Variables untuk CPM/Merchant data (10 kolom)
    DECLARE merchant_id VARCHAR(50);
    DECLARE pic_name VARCHAR(50);
    DECLARE mcc VARCHAR(10);
    DECLARE merchant_criteria VARCHAR(50);
    DECLARE province VARCHAR(50);
    DECLARE city VARCHAR(50);
    DECLARE district VARCHAR(50);
    DECLARE village VARCHAR(50);
    DECLARE kodepos VARCHAR(10);
    DECLARE approval_date VARCHAR(30);
    DECLARE v_approval_datetime DATETIME;
    
    -- Variables untuk Merchant Active QRIS (5 kolom)
    DECLARE tanggal_str VARCHAR(20);
    DECLARE city_active VARCHAR(50);
    DECLARE merchant_mcc VARCHAR(10);
    DECLARE criteria_active VARCHAR(50);
    DECLARE total_merchant VARCHAR(20);
    DECLARE v_total_merchant BIGINT;
    
    -- Variables untuk mapping
    DECLARE v_master_merchant_city VARCHAR(255);
    DECLARE v_merchant_category_name VARCHAR(255);
    
    -- Bersihkan kurung siku luar
    SET p_content = TRIM(BOTH '[]' FROM TRIM(p_content));
    
    -- Hitung jumlah baris
    SET row_count = (LENGTH(p_content) - LENGTH(REPLACE(p_content, '],[', ''))) / 3 + 1;
    
    IF row_count <= 0 THEN
        SIGNAL SQLSTATE '45000' SET MESSAGE_TEXT = 'No data rows to process';
    END IF;
    
    process_loop: WHILE i <= row_count DO
        -- Ambil satu baris
        IF i < row_count THEN
            SET end_pos = LOCATE('],[', p_content, start_pos);
            IF end_pos = 0 THEN LEAVE process_loop; END IF;
            SET current_row = SUBSTRING(p_content, start_pos, end_pos - start_pos);
            SET start_pos = end_pos + 3;
        ELSE
            SET current_row = SUBSTRING(p_content, start_pos);
        END IF;
        
        -- Bersihkan karakter ]] di akhir current_row (untuk baris terakhir)
        SET current_row = TRIM(TRAILING ']' FROM current_row);
        
        -- Skip baris kosong
        IF TRIM(current_row) = '' THEN
            SET i = i + 1;
            ITERATE process_loop;
        END IF;
        
        -- Hitung jumlah kolom dalam baris ini
        SET col_count = (LENGTH(current_row) - LENGTH(REPLACE(current_row, '","', ''))) + 1;
        
        -- Proses berdasarkan jumlah kolom dan tipe data
        IF p_dataType = 'autoreport_bi_merchant_qris' AND col_count >= 10 THEN
            -- === PROSES DATA MERCHANT BI QRIS (10 kolom) ===
            SET merchant_id = TRIM(BOTH '"' FROM SUBSTRING_INDEX(current_row, '","', 1));
            SET pic_name = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 2), '","', -1));
            SET mcc = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 3), '","', -1));
            SET merchant_criteria = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 4), '","', -1));
            SET province = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 5), '","', -1));
            SET city = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 6), '","', -1));
            SET district = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 7), '","', -1));
            SET village = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 8), '","', -1));
            SET kodepos = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 9), '","', -1));
            SET approval_date = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 10), '","', -1));
            
            -- Parse approval_date ke DATETIME
            SET v_approval_datetime = NULL;
            IF approval_date REGEXP '^[0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2}$' THEN
                SET v_approval_datetime = STR_TO_DATE(approval_date, '%Y-%m-%d %H:%i:%s');
            ELSEIF approval_date REGEXP '^[0-9]{2}-[0-9]{2}-[0-9]{4} [0-9]{2}:[0-9]{2}:[0-9]{2}$' THEN
                SET v_approval_datetime = STR_TO_DATE(approval_date, '%d-%m-%Y %H:%i:%s');
            END IF;
            
            -- Mapping merchant_category_name dari master_category
            SET v_merchant_category_name = NULL;
            SELECT mc.category
            INTO v_merchant_category_name
            FROM master_category mc
            WHERE mc.mcc = mcc
            LIMIT 1;
            
            -- Insert ke tabel master_data_merchant_cpm
            INSERT INTO master_data_merchant_cpm (
                merchant_id,
                nama_merchant_50,
                nama_merchant_25,
                merchant_category,
                merchant_criteria,
                merchant_category_name,
                status,
                periode_aktivasi,
                created_at
            ) VALUES (
                merchant_id,
                SUBSTRING(pic_name, 1, 50),
                SUBSTRING(pic_name, 1, 25),
                CAST(mcc AS UNSIGNED),
                merchant_criteria,
                v_merchant_category_name,
                'Active',
                v_approval_datetime,
                NOW()
            )
            ON DUPLICATE KEY UPDATE
                nama_merchant_50 = VALUES(nama_merchant_50),
                nama_merchant_25 = VALUES(nama_merchant_25),
                merchant_category = VALUES(merchant_category),
                merchant_criteria = VALUES(merchant_criteria),
                merchant_category_name = VALUES(merchant_category_name),
                periode_aktivasi = VALUES(periode_aktivasi);
        
        ELSEIF p_dataType = 'autoreport_merchant_active_qris' AND col_count >= 5 THEN
            -- === PROSES DATA MERCHANT ACTIVE QRIS (5 kolom) ===
            SET tanggal_str = TRIM(BOTH '"' FROM SUBSTRING_INDEX(current_row, '","', 1));
            SET city_active = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 2), '","', -1));
            SET merchant_mcc = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 3), '","', -1));
            SET criteria_active = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 4), '","', -1));
            SET total_merchant = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 5), '","', -1));
            
            -- Deteksi format tanggal
            SET v_transaction_date = NULL;
            IF tanggal_str REGEXP '^[0-9]{4}-[0-9]{2}-[0-9]{2}$' THEN
                SET v_transaction_date = STR_TO_DATE(tanggal_str, '%Y-%m-%d');
            ELSEIF tanggal_str REGEXP '^[0-9]{2}-[0-9]{2}-[0-9]{4}$' THEN
                SET v_transaction_date = STR_TO_DATE(tanggal_str, '%d-%m-%Y');
            ELSE
                SET i = i + 1;
                ITERATE process_loop;
            END IF;
            
            IF v_transaction_date IS NULL THEN
                SET i = i + 1;
                ITERATE process_loop;
            END IF;
            
            -- Parse total merchant
            SET v_total_merchant = IF(total_merchant = '' OR total_merchant IS NULL, 0, CAST(REPLACE(total_merchant, ',', '') AS UNSIGNED));
            
            -- Mapping master_merchant_city dari master_merchant_city
            SET v_master_merchant_city = NULL;
            SELECT mmc.master_merchant_city
            INTO v_master_merchant_city
            FROM master_merchant_city mmc
            WHERE mmc.transaction_merchant_city = city_active
            LIMIT 1;
            
            -- Mapping merchant_category_name dari master_category
            SET v_merchant_category_name = NULL;
            SELECT mc.category
            INTO v_merchant_category_name
            FROM master_category mc
            WHERE mc.mcc = merchant_mcc
            LIMIT 1;
            
            -- Insert ke tabel master_data
            INSERT INTO master_data (
                transaction_date,
                merchant_city,
                merchant_category,
                merchant_criteria,
                amount_trx,
                merchant_category_name,
                nama_kab_kota,
                master_merchant_city,
                created_at,
                master_data_type
            ) VALUES (
                v_transaction_date,
                city_active,
                CAST(merchant_mcc AS UNSIGNED),
                criteria_active,
                v_total_merchant,
                v_merchant_category_name,
                city_active,
                v_master_merchant_city,
                NOW(),
                p_dataType
            );
            
        ELSEIF col_count >= 7 THEN
            -- === PROSES DATA MPM (7 kolom) ===
            SET date_str = TRIM(BOTH '"' FROM SUBSTRING_INDEX(current_row, '","', 1));
            SET col2 = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 2), '","', -1));
            SET col3 = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 3), '","', -1));
            SET col4 = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 4), '","', -1));
            SET col5 = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 5), '","', -1));
            SET col6 = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 6), '","', -1));
            SET col7 = TRIM(BOTH '"' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(current_row, '","', 7), '","', -1));
            
            -- Deteksi format tanggal
            IF date_str REGEXP '^[0-9]{4}-[0-9]{2}-[0-9]{2}$' THEN
                SET v_transaction_date = STR_TO_DATE(date_str, '%Y-%m-%d');
            ELSEIF date_str REGEXP '^[0-9]{2}-[0-9]{2}-[0-9]{4}$' THEN
                SET v_transaction_date = STR_TO_DATE(date_str, '%d-%m-%Y');
            ELSE
                SET i = i + 1;
                ITERATE process_loop;
            END IF;
            
            IF v_transaction_date IS NULL THEN
                SET i = i + 1;
                ITERATE process_loop;
            END IF;
            
            -- Parse angka
            SET v_jumlah_trx = IF(col6 = '' OR col6 IS NULL, 0, CAST(col6 AS UNSIGNED));
            SET v_amount_trx = IF(col7 = '' OR col7 IS NULL, 0, CAST(REPLACE(col7, ',', '') AS DECIMAL(18,0)));
            
            -- Mapping master_merchant_city dari master_merchant_city
            SET v_master_merchant_city = NULL;
            SELECT mmc.master_merchant_city
            INTO v_master_merchant_city
            FROM master_merchant_city mmc
            WHERE mmc.transaction_merchant_city = col2
            LIMIT 1;
            
            -- Mapping merchant_category_name dari master_category
            SET v_merchant_category_name = NULL;
            SELECT mc.category
            INTO v_merchant_category_name
            FROM master_category mc
            WHERE mc.mcc = col4
            LIMIT 1;
            
            -- Insert ke tabel master_data
            INSERT INTO master_data (
                transaction_date,
                merchant_city,
                keterangan,
                merchant_category,
                merchant_criteria,
                jumlah_trx,
                amount_trx,
                merchant_category_name,
                nama_kab_kota,
                master_merchant_city,
                created_at,
                master_data_type
            ) VALUES (
                v_transaction_date,
                col2,
                col3,
                col4,
                col5,
                v_jumlah_trx,
                v_amount_trx,
                v_merchant_category_name,
                col2,
                v_master_merchant_city,
                NOW(),
                p_dataType
            );
        END IF;
        
        SET i = i + 1;
    END WHILE process_loop;
    
END


1. Create data table
   Source data type: Text File
   Text File: "C:\\Users\\FANI\\PT ELISTEC INFORMATIKA UTAMA\\Ahmad Zidni Zainul Ikhsan - Phase 2\\Module 2 - Regulatory\\Laporan QRIS\\config_email.csv"
   File Encoding: UTF-8
   Generation settings: 
      Automatically detect column type: enable
      Use the first row as the table header: enable
      format: CSV (Separated by comma)
      column separator: [semicolon]
     save result: configMail

2. Configure MySQL Database  (Pindah ke atas, cukup sekali)
   Database address: "localhost"
   port: 3306
   username: "root"
   password: 12345
   database name: "regulatorydev"
   Database configuration object: regulatorydbdev

   console log: "Starting process to retrieve email attachments"

3. Traverse data table  ← LOOP UTAMA (semua proses di dalam sini)
   Data table: configMail
   Traversal mode: Traverse by row
   row number: _row_index
   row data: _row_value

   3.1 Get Array Element of Specified Type
      Array: _row_value
      Target element: 0
      Asserted element type: string
      Result: getKeyType   
           

   3.2 Get Array Element of Specified Type
      Array: _row_value
      Target element: 1
      Asserted element type: string
      Result: getValue

   3.3 Split String
      String to be split: getValue
      Separator: Customize
      Custom: "|"
      Filter empty: true
      save result: splitValue

   3.4 Get Array Elements
      Array: splitValue
      Target element: 0
      Result: subject

   3.5 Get Array Elements
      Array: splitValue
      Target element: 1
      Result: sender

   3.6 Get Array Elements
      Array: splitValue
      Target element: 2
      Result: timeFrom

   3.7 Get Array Elements
      Array: splitValue
      Target element: 3
      Result: timeTo

   3.8 Get Current Date and Time
      current datetime: getDateTime

   3.9 Date To String
      Datetime Object: getDateTime
      Output string format: YYYY-MM-DD
      Conversion Result: todayString

   3.10 Get Mails
      selected_mail: false
      matching_method: Precise
      include_subfolders: false
      limit_emails_to_first: 10
      unread_mail: true
      attachments_only: true
      login_account: "rpa@finnet.co.id"
      folders: ["inbox"]
      date_from: todayString + " " + timeFrom
      date_to: todayString + " " + timeTo
      sender_mails: [sender]
      mail_subject: subject
      Mail_ID_Array: rpaMailArray (array<any>)


      console log : _row_index+". "+"Get Email: "+subject

   3.11 Traverse array  (jika ada email)
      Array: rpaMailArray
      Element’s subscript: row_index
      Elements in array: currentMailID


      3.11.1 Get_Email_By_ID
         Mail ID: currentMailID
         Mail Object: currentMailObject

      3.11.2 Save Mail Attachments
         Mail ID: currentMailID
         exclude_inline_attachments: false
         save_to_folder: "C:\\RPA\\FinnetDev\\attachments\\"
         filter_by_file_name: ["*.xlsx", "*.xls"]
         overwrite_existing: true
         file array: currentAttachments

        console log: "Save Attachments: " + currentAtachments[0]
      
      3.11.3 Get Array Length  ← PINDAH KE SINI (SUB dari 3.11)
         Array: currentAttachments
         Array length: attachmentsCount
      
      3.11.4 If - Conditional judgment:  ← PINDAH KE SINI (SUB dari 3.11)
         IF attachmentsCount > 0
            
            3.11.4.1 Traverse Array
               Array: currentAttachments
               Element's subscript: file_idx
               Elements in array: currentFilePath

               3.11.4.1.1 Assign Value To Variable
                  Value: currentFilePath
                  Variabel name: currentFilePathToString

               3.11.4.1.2 Open Excel Workbook
                  File path: currentFilePath
                  When the file does not exist: Do not automatically
                  Open mode: automatic detection
                  Is it visible: No
                  Excel File: currentExcelOBJ

               3.11.4.1.3 Read data within range
                  Excel file object: currentExcelOBJ
                  Sheet name/serial number: 1
                  Read data within range: whole worksheet
                  Read options: Real Value
                  Return type: DataTable
                  Set table header: Yes
                  Read result: excelDataResult

                  console log: "Open & Read Excel " + getKeyType

               3.11.4.1.4 Convert to Array
                  Data Table: excelDataResult
                  Include Table Header: No
                  Conversion result: excelDataNew
               
               3.11.4.1.5 Variable To String
                  orginal variabel: excelDataNew
                  convertion result: excelDataString

               3.11.4.1.6 Execute SQL Statements
                  SQL Statements: ["CALL InsertQrisRegulatoryData('" + getKeyType + "', '" + excelDataString + "')"]
                  Database config: regulatorydbdev
                  output format: Recordset
                  timeout (second): 300
                  Query result: queryInsert
                  
                  console log: "Success Save Attachments to database"

               3.11.4.1.6 Close Excel Workbook
                  Excel file object: currentExcelOBJ
                  Closing Method: Close Without Saving
      
            3.11.4.2 Mark Email as Read 
               Mail ID: currentMailID
               Mark as: Read

    4. Console Log: Completed Successfully

    5. ⭐ Archive Files to DONE Folder
        Console Log: "Starting archive process..."

        5.1 Assign Value To Variable
            Value: base_path + "\\attachments\\"
            Variable name: attachmentsFolder

        5.2 Assign Value To Variable
            Value: "C:\\RPA\\INNETDEV\\Usecase 1- Qris Report\\Autoreport Download\\DONE\\"
            Variable name: baseDonePath

        5.3 Get the files or folders in the directory
            Directory path: attachmentsFolder
            Filter: "*.xlsx"
            Include subdirectories: No
            Result: allFilesInAttachments (array)

        5.4 Get Array Length
            Array: allFilesInAttachments
            Array length: totalFilesToMove

        5.5 Console Log: "Found " + totalFilesToMove + " file(s) to archive"

        5.6 If - Conditional judgment:
            IF totalFilesToMove > 0
            THEN:
                5.6.1 Traverse Array
                    Array: allFilesInAttachments
                    Element's subscript: archive_idx
                    Elements in array: filePathToMove

                    5.6.1.1 Assign Value To Variable
                    Value: filePathToMove
                    Variable name: filePathToMoveString

                    5.6.1.2 Split String (Extract Filename)
                    String to be split: filePathToMoveString
                    Separator: Custom separator
                    Custom: "\\"
                    Filter empty: true
                    Save result: pathParts

                    5.6.1.3 Get Array Length
                    Array: pathParts
                    Array length: pathPartsCount

                    5.6.1.4 Get Array Elements
                    Array: pathParts
                    Target Element Subscript: pathPartsCount - 1
                    Result: extractedFileName

                    5.6.1.5 Split Sting
                    String to be split: extractedFileName
                    Separator: Custom separator
                    Custom: "_"
                    Filter empty: true
                    Save result: fileParts (array)

                    5.6.1.6 Get Array Length
                    Array: fileParts
                    Array length: filePartsCount

                    5.6.1.7 If - Conditional judgment (Determine Target Folder):
                    (fileParts[filePartsCount - 3] == "off" || fileParts[filePartsCount - 3] == "on") && fileParts[filePartsCount - 2] == "us"
                    THEN:
                    Assign Value To Variable
                     Value: baseDonePath + "MPM\\"
                     Variable name: targetFolder

                    ELSE IF fileParts[filePartsCount-2]=="cpm"
                    THEN:
                        Assign Value To Variable
                            Value: baseDonePath + "CPM\\"
                            Variable name: targetFolder

                    ELSE IF fileParts[1] == "bi" && fileParts[2] == "merchant"
                    THEN:
                        Assign Value To Variable
                            Value: baseDonePath + "Merchant\\BI Merchant\\"
                            Variable name: targetFolder

                    ELSE IF fileParts[1] == "merchant" && fileParts[2] == "active"
                    THEN:
                        Assign Value To Variable
                            Value: baseDonePath + "Merchant\\BI Merchant\\"
                            Variable name: targetFolder

                    ELSE:
                        Assign Value To Variable
                            Value: baseDonePath + "Others\\"
                            Variable name: targetFolder
            
            5.7 Move files to the specified directory
               Original File Path: attachmentsFolder + filePathToMoveString
               Folder Path: targetFolder

DONE/
├── MPM/
│   ├── Autoreport_qris_acquirer_off_us.xlsx
│   ├── Autoreport_qris_issuer_off_us.xlsx
│   └── Autoreport_qris_issuer_on_us.xlsx
│
├── CPM/
│   ├── Autoreport_qris_acquirer_off_us_cpm.xlsx
│   ├── Autoreport_qris_issuer_off_us_cpm.xlsx
│   └── Autoreport_qris_issuer_on_us_cpm.xlsx
│
└── Merchant/
    ├── Merchant Active/
    │   └── Autoreport_merchant_active_qris.xlsx
    │
    └── BI Merchant/
        └── Autoreport_bi_merchant_qris.xlsx


