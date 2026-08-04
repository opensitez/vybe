*> vybe-test: cobol/builtin_library_features/xml_parse_feature_compiles
*> origin: languages/cobol/tests/cobol/test_builtin_library_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-XML PIC X(200) VALUE "<a>1</a>".
PROCEDURE DIVISION.
    XML PARSE WS-XML PROCESSING PROCEDURE X-HANDLER.
    STOP RUN.
X-HANDLER SECTION.
    DISPLAY "TAG".

