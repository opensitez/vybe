*> vybe-test: cobol/category_inspect/test_inspect_before_after
*> origin: languages/cobol/tests/cobol/test_category_inspect.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. INSPECT-BA.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 STR PIC X(20) VALUE "ABC*DEF*GHI".
       01 CNT PIC 9(2) VALUE 0.
       PROCEDURE DIVISION.
           INSPECT STR TALLYING CNT FOR CHARACTERS
              BEFORE INITIAL "*".
           DISPLAY CNT.
    MOVE SPACES TO WS-VYBE-L
    STRING CNT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "03"
        DISPLAY "FAIL: want [03] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

