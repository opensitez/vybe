*> vybe-test: cobol/category_data_division_renames/test_renames_parse_numeric_subtract
*> origin: languages/cobol/tests/cobol/test_category_data_division_renames.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 N. 05 A PIC 99 VALUE 12. 05 B PIC 99 VALUE 34. 66 NUM RENAMES A THRU B. PROCEDURE DIVISION. SUBTRACT 34 FROM NUM DISPLAY NUM.
    MOVE SPACES TO WS-VYBE-L
    STRING NUM DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "1200"
        DISPLAY "FAIL: want [1200] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

