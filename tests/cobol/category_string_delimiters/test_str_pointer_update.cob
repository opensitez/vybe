*> vybe-test: cobol/category_string_delimiters/test_str_pointer_update
*> origin: languages/cobol/tests/cobol/test_category_string_delimiters.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 S1 PIC X(2) VALUE 'AB'. 01 R PIC X(5). 01 P PIC 9 VALUE 1. PROCEDURE DIVISION. STRING S1 DELIMITED BY SIZE INTO R WITH POINTER P. DISPLAY P.
    MOVE SPACES TO WS-VYBE-L
    STRING P DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "3"
        DISPLAY "FAIL: want [3] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

