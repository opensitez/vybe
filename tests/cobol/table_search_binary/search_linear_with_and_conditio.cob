*> vybe-test: cobol/table_search_binary/search_linear_with_and_condition
*> origin: languages/cobol/tests/cobol/test_table_search_binary.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E OCCURS 10 TIMES INDEXED BY IX.
      10 CODE PIC X(2).
      10 STATUS PIC X VALUE "A".
PROCEDURE DIVISION.
    SET IX TO 1.
    SEARCH E
        AT END DISPLAY "NOT FOUND"
        WHEN CODE(IX) = "AB" AND STATUS(IX) = "A"
            DISPLAY "FOUND"
    END-SEARCH.
    STOP RUN.

