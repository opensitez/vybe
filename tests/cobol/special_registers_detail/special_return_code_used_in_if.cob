*> vybe-test: cobol/special_registers_detail/special_return_code_used_in_if
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 RETURN-CODE PIC 9(4) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF RETURN-CODE = 0
        DISPLAY "SUCCESS"
    ELSE
        DISPLAY "FAILURE"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "SUCCESS" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "SUCCESS"
        DISPLAY "FAIL: want [SUCCESS] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

