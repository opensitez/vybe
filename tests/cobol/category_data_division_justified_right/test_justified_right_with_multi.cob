*> vybe-test: cobol/category_data_division_justified_right/test_justified_right_with_multiple_assignments
*> origin: languages/cobol/tests/cobol/test_category_data_division_justified_right.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 A PIC X(4) JUSTIFIED RIGHT. 01 B PIC X(4) JUSTIFIED RIGHT. PROCEDURE DIVISION. MOVE 'AB' TO A. MOVE 'Z' TO B. DISPLAY '[' A ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE A DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[  AB]"
        DISPLAY "FAIL: want [[  AB]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. DISPLAY '[' B ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE B DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[   Z]"
        DISPLAY "FAIL: want [[   Z]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

