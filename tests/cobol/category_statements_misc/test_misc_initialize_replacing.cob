*> vybe-test: cobol/category_statements_misc/test_misc_initialize_replacing
*> origin: languages/cobol/tests/cobol/test_category_statements_misc.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. MISC-INIT-REPL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 GRP.
          05 FLD-1 PIC X(3).
          05 FLD-2 PIC 9(3).
       PROCEDURE DIVISION.
           INITIALIZE GRP REPLACING ALPHANUMERIC BY "A" NUMERIC BY 9.
           DISPLAY FLD-1 " " FLD-2.
    MOVE SPACES TO WS-VYBE-L
    STRING FLD-1 DELIMITED SIZE " " DELIMITED SIZE FLD-2 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "AAA 999"
        DISPLAY "FAIL: want [AAA 999] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

