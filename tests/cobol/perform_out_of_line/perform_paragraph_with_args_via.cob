*> vybe-test: cobol/perform_out_of_line/perform_paragraph_with_args_via_working_storage
*> origin: languages/cobol/tests/cobol/test_perform_out_of_line.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9(3) VALUE 7.
01 Y PIC 9(3) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    PERFORM DOUBLE.
    DISPLAY Y.
    MOVE SPACES TO WS-VYBE-L
    STRING Y DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "14"
        DISPLAY "FAIL: want [14] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.
DOUBLE.
    MULTIPLY X BY 2 GIVING Y.
    STOP RUN.

