*> vybe-test: cobol/category_data_division_synchronized/test_sync_multiple_fields_display
*> origin: languages/cobol/tests/cobol/test_category_data_division_synchronized.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 F1 PIC X(3) VALUE 'A' SYNCHRONIZED. 01 F2 PIC X(3) VALUE 'B' SYNCHRONIZED. PROCEDURE DIVISION. DISPLAY '[' F1 ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE F1 DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[A  ]"
        DISPLAY "FAIL: want [[A  ]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. DISPLAY '[' F2 ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE F2 DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[B  ]"
        DISPLAY "FAIL: want [[B  ]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

