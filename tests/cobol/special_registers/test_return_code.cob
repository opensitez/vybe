*> vybe-test: cobol/special_registers/test_return_code
*> origin: languages/cobol/tests/cobol/test_special_registers.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.

    MOVE 8 TO RETURN-CODE.
    IF RETURN-CODE NOT = 0
        DISPLAY "ERROR"
    END-IF.
    STOP RUN.

