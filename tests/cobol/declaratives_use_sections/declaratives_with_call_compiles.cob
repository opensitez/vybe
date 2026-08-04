*> vybe-test: cobol/declaratives_use_sections/declaratives_with_call_compiles
*> origin: languages/cobol/tests/cobol/test_declaratives_use_sections.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
DECLARATIVES.
D2 SECTION.
    USE AFTER STANDARD ERROR PROCEDURE ON WS-FILE.
END DECLARATIVES.
M2 SECTION.
    CALL "WORK".
    STOP RUN.

