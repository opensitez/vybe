*> vybe-test: cobol/builtin_libraries_common/xml_parse_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_libraries_common.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(100).
PROCEDURE DIVISION.
    XML PARSE X PROCESSING PROCEDURE H.
    STOP RUN.
H SECTION.
    DISPLAY "E".

