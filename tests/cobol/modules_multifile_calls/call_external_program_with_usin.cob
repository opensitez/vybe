*> vybe-test: cobol/modules_multifile_calls/call_external_program_with_using_compiles
*> origin: languages/cobol/tests/cobol/test_modules_multifile_calls.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-B.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(10).
PROCEDURE DIVISION.
    CALL "SUB-B" USING WS-A.
    STOP RUN.

