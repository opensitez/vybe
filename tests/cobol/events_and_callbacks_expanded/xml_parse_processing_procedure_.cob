*> vybe-test: cobol/events_and_callbacks_expanded/xml_parse_processing_procedure_compiles
*> origin: languages/cobol/tests/cobol/test_events_and_callbacks_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(200) VALUE "<a>1</a>".
PROCEDURE DIVISION.
    XML PARSE X PROCESSING PROCEDURE P-H.
    STOP RUN.
P-H SECTION.
    DISPLAY "H".

