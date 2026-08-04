*> vybe-test: cobol/modules_multifile_calls/call_external_program_by_content_compiles
*> origin: languages/cobol/tests/cobol/test_modules_multifile_calls.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-D.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(5) VALUE 10.
PROCEDURE DIVISION.
    CALL "SUB-D" USING BY CONTENT WS-A.
    STOP RUN.

