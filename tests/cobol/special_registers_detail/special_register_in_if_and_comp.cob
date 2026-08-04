*> vybe-test: cobol/special_registers_detail/special_register_in_if_and_compute
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 RETURN-CODE PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE RETURN-CODE = 4 + 4.
    STOP RUN.

