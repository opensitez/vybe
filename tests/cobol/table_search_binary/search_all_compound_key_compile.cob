*> vybe-test: cobol/table_search_binary/search_all_compound_key_compiles
*> origin: languages/cobol/tests/cobol/test_table_search_binary.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 COMPOUND-T.
   05 CT OCCURS 20 TIMES ASCENDING KEY CT-KEY1 CT-KEY2 INDEXED BY CI.
      10 CT-KEY1 PIC X(4).
      10 CT-KEY2 PIC 9(4).
      10 CT-DATA PIC X(20).
PROCEDURE DIVISION.
    SEARCH ALL CT
        AT END DISPLAY "NOT FOUND"
        WHEN CT-KEY1(CI) = "ABCD" AND CT-KEY2(CI) = 1001
            DISPLAY CT-DATA(CI)
    END-SEARCH.
    STOP RUN.

