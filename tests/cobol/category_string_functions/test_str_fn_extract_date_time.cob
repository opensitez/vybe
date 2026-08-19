*> vybe-test: cobol/category_string_functions/test_str_fn_extract_date_time
*> origin: languages/cobol/tests/cobol/test_category_string_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 D PIC X(8) VALUE '20230101'. PROCEDURE DIVISION. DISPLAY FUNCTION EXTRACT-DATE-TIME(D '%Y').
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION EXTRACT-DATE-TIME(D, '%Y') DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "2023"
        DISPLAY "FAIL: want [2023] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

