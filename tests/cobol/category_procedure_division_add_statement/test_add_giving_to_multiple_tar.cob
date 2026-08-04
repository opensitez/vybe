*> vybe-test: cobol/category_procedure_division_add_statement/test_add_giving_to_multiple_targets_runtime
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_add_statement.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 A PIC 9 VALUE 10. 01 B PIC 9 VALUE 20. 01 R1 PIC 99. 01 R2 PIC 99. PROCEDURE DIVISION. ADD A B GIVING R1 R2. DISPLAY R1.
    MOVE SPACES TO WS-VYBE-L
    STRING R1 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "30"
        DISPLAY "FAIL: want [30] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. DISPLAY R2.
    MOVE SPACES TO WS-VYBE-L
    STRING R2 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "30"
        DISPLAY "FAIL: want [30] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

