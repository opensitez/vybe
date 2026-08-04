*> vybe-test: cobol/special_registers_detail/special_return_code_display
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 RETURN-CODE PIC 9(4) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 0 TO RETURN-CODE.
    DISPLAY RETURN-CODE.
    MOVE SPACES TO WS-VYBE-L
    STRING RETURN-CODE DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0000"
        DISPLAY "FAIL: want [0000] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

