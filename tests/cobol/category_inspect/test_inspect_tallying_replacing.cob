*> vybe-test: cobol/category_inspect/test_inspect_tallying_replacing_combined
*> origin: languages/cobol/tests/cobol/test_category_inspect.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. INSPECT-COMB.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 STR PIC X(10) VALUE "  123   ".
       01 CNT PIC 9(2) VALUE 0.
       PROCEDURE DIVISION.
           INSPECT STR
              TALLYING CNT FOR LEADING " "
              REPLACING LEADING " " BY "0"
                        ALL " " BY "X".
           DISPLAY STR " " CNT.
    MOVE SPACES TO WS-VYBE-L
    STRING STR DELIMITED SIZE " " DELIMITED SIZE CNT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "00123XXX   02"
        DISPLAY "FAIL: want [00123XXX   02] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

