*> vybe-test: cobol/data_group_level/data_group_with_value_low_values
*> origin: languages/cobol/tests/cobol/test_data_group_level.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 REC.
   05 MARKER PIC X(4) VALUE LOW-VALUES.
PROCEDURE DIVISION.
    DISPLAY REC.
    STOP RUN.

