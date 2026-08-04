*> vybe-test: cobol/data_group_level/data_group_with_comp_field
*> origin: languages/cobol/tests/cobol/test_data_group_level.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 REC.
   05 TEXT PIC X(10) VALUE "HELLO".
   05 BINARY-NUM PIC 9(8) COMP VALUE 0.
PROCEDURE DIVISION.
    ADD 1 TO BINARY-NUM.
    STOP RUN.

