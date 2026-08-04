*> vybe-test: cobol/special_registers_detail/return_code_subtract
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 RETURN-CODE PIC 9(4) VALUE 8.
PROCEDURE DIVISION.
    SUBTRACT 4 FROM RETURN-CODE.
    STOP RUN.

