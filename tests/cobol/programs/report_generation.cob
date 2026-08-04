*> vybe-test: cobol/programs/report_generation
*> origin: languages/cobol/tests/cobol/test_programs.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. REPORT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-HEADER PIC X(60).
01 WS-LINE   PIC X(60).
01 WS-TOTAL  PIC 9(8)V99 VALUE 0.
01 WS-I      PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    STRING "Product" DELIMITED BY SIZE
           "    " DELIMITED BY SIZE
           "Price" DELIMITED BY SIZE
           INTO WS-HEADER.
    DISPLAY WS-HEADER.
    DISPLAY "----------------------------".
    ADD 29.99 TO WS-TOTAL.
    ADD 49.99 TO WS-TOTAL.
    ADD 19.99 TO WS-TOTAL.
    DISPLAY "Total: " WS-TOTAL.
    STOP RUN.

