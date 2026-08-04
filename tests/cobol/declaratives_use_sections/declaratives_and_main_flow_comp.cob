*> vybe-test: cobol/declaratives_use_sections/declaratives_and_main_flow_compiles
*> origin: languages/cobol/tests/cobol/test_declaratives_use_sections.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
DECLARATIVES.
D1 SECTION.
    USE AFTER STANDARD ERROR PROCEDURE ON WS-FILE.
END DECLARATIVES.
M1 SECTION.
    DISPLAY "MAIN".
    STOP RUN.

