*> vybe-test: cobol/special_registers_detail/special_return_code_move_nonzero
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 RETURN-CODE PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    MOVE 8 TO RETURN-CODE.
    STOP RUN.

