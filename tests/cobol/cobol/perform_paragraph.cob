*> vybe-test: cobol/cobol/perform_paragraph
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. PPARA.
PROCEDURE DIVISION.
    PERFORM GREET-PARA.
    STOP RUN.
GREET-PARA.
    DISPLAY "Hello from paragraph".

