*> vybe-test: cobol/category_data_division/test_data_level_88_multiple
*> origin: languages/cobol/tests/cobol/test_category_data_division.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. DATA-88.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 STATUS-CODE PIC X.
          88 IS-VALID VALUE "A" "B" "C".
       PROCEDURE DIVISION.
           MOVE "B" TO STATUS-CODE.
           IF IS-VALID
              DISPLAY "VALID"
           END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "VALID" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "VALID"
        DISPLAY "FAIL: want [VALID] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

