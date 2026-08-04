*> vybe-test: cobol/final_features/copy_replacing_multi
*> origin: languages/cobol/tests/cobol/test_final_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    COPY RECORD-DEF REPLACING "OLD" BY "NEW" "FIELD1" BY "FIELD2".
    STOP RUN.

