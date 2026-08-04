*> vybe-test: cobol/table_search_binary/search_all_numeric_at_end
*> origin: languages/cobol/tests/cobol/test_table_search_binary.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 NUMS.
   05 NUM-ENTRY OCCURS 50 TIMES ASCENDING KEY NUM-VAL INDEXED BY NI.
      10 NUM-VAL PIC 9(6).
PROCEDURE DIVISION.
    SEARCH ALL NUM-ENTRY
        AT END
            DISPLAY "VALUE NOT IN TABLE"
        WHEN NUM-VAL(NI) = 999999
            DISPLAY "FOUND MAX"
    END-SEARCH.
    STOP RUN.

