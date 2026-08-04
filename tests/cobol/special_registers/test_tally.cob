*> vybe-test: cobol/special_registers/test_tally
*> origin: languages/cobol/tests/cobol/test_special_registers.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STR PIC X(10) VALUE "A B C D".
PROCEDURE DIVISION.

    INSPECT WS-STR TALLYING TALLY FOR ALL " ".
    STOP RUN.

