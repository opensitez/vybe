*> vybe-test: cobol/cobol2023_nested/exit_section
*> origin: languages/cobol/tests/cobol/test_cobol2023_nested.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    PERFORM WORK-PARA.
    DISPLAY "Done".
    STOP RUN.
WORK-PARA.
    DISPLAY "Working".
    EXIT SECTION.
    DISPLAY "Never reached".

