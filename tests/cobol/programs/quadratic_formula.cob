*> vybe-test: cobol/programs/quadratic_formula
*> origin: languages/cobol/tests/cobol/test_programs.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. QUADRATIC.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A     PIC S9(5)V99 VALUE 1.
01 WS-B     PIC S9(5)V99 VALUE -5.
01 WS-C     PIC S9(5)V99 VALUE 6.
01 WS-DISC  PIC S9(10)V99 VALUE 0.
01 WS-ROOT1 PIC S9(10)V99 VALUE 0.
01 WS-ROOT2 PIC S9(10)V99 VALUE 0.
PROCEDURE DIVISION.
    COMPUTE WS-DISC = WS-B ** 2 - 4 * WS-A * WS-C.
    COMPUTE WS-ROOT1 = (-1 * WS-B + FUNCTION SQRT(WS-DISC))
                       / (2 * WS-A).
    COMPUTE WS-ROOT2 = (-1 * WS-B - FUNCTION SQRT(WS-DISC))
                       / (2 * WS-A).
    DISPLAY "Root 1: " WS-ROOT1.
    DISPLAY "Root 2: " WS-ROOT2.
    STOP RUN.

