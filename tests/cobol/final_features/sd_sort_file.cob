*> vybe-test: cobol/final_features/sd_sort_file
*> origin: languages/cobol/tests/cobol/test_final_features.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. SDTEST.
DATA DIVISION.
FILE SECTION.
SD SORT-FILE.
01 SORT-RECORD.
   05 SORT-KEY PIC 9(5).
   05 SORT-DATA PIC X(75).
WORKING-STORAGE SECTION.
01 WS-REC PIC X(80).
PROCEDURE DIVISION.
    DISPLAY "SD Test".
    STOP RUN.

