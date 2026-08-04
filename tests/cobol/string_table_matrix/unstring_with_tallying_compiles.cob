*> vybe-test: cobol/string_table_matrix/unstring_with_tallying_compiles
*> origin: languages/cobol/tests/cobol/test_string_table_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC X(12) VALUE "A,B,C".
01 F1 PIC X(2).
01 F2 PIC X(2).
01 F3 PIC X(2).
01 T PIC 9 VALUE 0.
PROCEDURE DIVISION.
    UNSTRING SRC DELIMITED BY "," INTO F1 F2 F3 TALLYING IN T.
    STOP RUN.

