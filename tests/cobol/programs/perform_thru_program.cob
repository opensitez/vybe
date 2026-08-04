*> vybe-test: cobol/programs/perform_thru_program
*> origin: languages/cobol/tests/cobol/test_programs.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. PTHRU.
PROCEDURE DIVISION.
    PERFORM STEP-1 THRU STEP-3.
    DISPLAY "All steps complete".
    STOP RUN.
STEP-1.
    DISPLAY "Step 1: Initialize".
STEP-2.
    DISPLAY "Step 2: Process".
STEP-3.
    DISPLAY "Step 3: Finalize".

