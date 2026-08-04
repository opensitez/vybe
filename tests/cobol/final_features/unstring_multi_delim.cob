*> vybe-test: cobol/final_features/unstring_multi_delim
*> origin: languages/cobol/tests/cobol/test_final_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SRC PIC X(30) VALUE "A,B;C".
01 F1 PIC X(10).
01 F2 PIC X(10).
01 F3 PIC X(10).
PROCEDURE DIVISION.
    UNSTRING WS-SRC DELIMITED BY "," OR ";" INTO F1 F2 F3.
    STOP RUN.

