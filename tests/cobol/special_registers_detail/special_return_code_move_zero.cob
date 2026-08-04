*> vybe-test: cobol/special_registers_detail/special_return_code_move_zero
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 RETURN-CODE PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    MOVE 0 TO RETURN-CODE.
    STOP RUN.

