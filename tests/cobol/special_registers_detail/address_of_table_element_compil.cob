*> vybe-test: cobol/special_registers_detail/address_of_table_element_compiles
*> origin: languages/cobol/tests/cobol/test_special_registers_detail.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC X(5) OCCURS 10 TIMES.
01 P USAGE POINTER.
PROCEDURE DIVISION.
    SET P TO ADDRESS OF E(1).
    STOP RUN.

