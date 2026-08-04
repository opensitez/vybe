*> vybe-test: cobol/cobol/paragraph_perform_thru
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. PARATHRU.
PROCEDURE DIVISION.
    PERFORM INIT-PARA.
    PERFORM PROCESS-PARA.
    STOP RUN.
INIT-PARA.
    DISPLAY "Initializing".
PROCESS-PARA.
    DISPLAY "Processing".

