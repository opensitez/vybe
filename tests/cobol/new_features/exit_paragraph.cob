*> vybe-test: cobol/new_features/exit_paragraph
*> origin: languages/cobol/tests/cobol/test_new_features.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    PERFORM MY-PARA.
    STOP RUN.
MY-PARA.
    DISPLAY "Start".
    EXIT PARAGRAPH.
    DISPLAY "Never".

