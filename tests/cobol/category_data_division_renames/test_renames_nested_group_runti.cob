*> vybe-test: cobol/category_data_division_renames/test_renames_nested_group_runtime
*> origin: languages/cobol/tests/cobol/test_category_data_division_renames.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 OUTER. 05 A PIC X(2) VALUE 'AA'. 05 INNER. 10 B PIC X VALUE 'B'. 10 C PIC X VALUE 'C'. 66 RENAME-GRP RENAMES B THRU C. PROCEDURE DIVISION. DISPLAY RENAME-GRP.
    MOVE SPACES TO WS-VYBE-L
    STRING RENAME-GRP DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "BC"
        DISPLAY "FAIL: want [BC] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

