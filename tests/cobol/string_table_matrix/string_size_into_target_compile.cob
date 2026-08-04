*> vybe-test: cobol/string_table_matrix/string_size_into_target_compiles
*> origin: languages/cobol/tests/cobol/test_string_table_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(3) VALUE "ONE".
01 B PIC X(3) VALUE "TWO".
01 R PIC X(10).
PROCEDURE DIVISION.
    STRING A DELIMITED BY SIZE B DELIMITED BY SIZE INTO R.
    STOP RUN.

