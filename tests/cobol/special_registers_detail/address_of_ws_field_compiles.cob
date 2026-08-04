*> vybe-test: cobol/special_registers_detail/address_of_ws_field_compiles
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(4) VALUE 0.
01 PTR USAGE POINTER.
PROCEDURE DIVISION.
    SET PTR TO ADDRESS OF N.
    STOP RUN.

