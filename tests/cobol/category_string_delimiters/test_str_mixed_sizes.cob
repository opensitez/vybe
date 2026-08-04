*> vybe-test: cobol/category_string_delimiters/test_str_mixed_sizes
*> origin: languages/cobol/tests/cobol/test_category_string_delimiters.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 S1 PIC X VALUE 'A'. 01 S2 PIC X(2) VALUE 'BC'. 01 R PIC X(4). PROCEDURE DIVISION. STRING S1 S2 DELIMITED BY SIZE INTO R. DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABC "
        DISPLAY "FAIL: want [ABC ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

