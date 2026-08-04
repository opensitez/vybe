*> vybe-test: cobol/builtin_libraries_common/json_parse_into_group_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_libraries_common.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 J PIC X(100).
01 R.
   05 N PIC X(10).
PROCEDURE DIVISION.
    JSON PARSE J INTO R.
    STOP RUN.

