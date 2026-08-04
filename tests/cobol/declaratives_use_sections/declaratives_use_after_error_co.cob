*> vybe-test: cobol/declaratives_use_sections/declaratives_use_after_error_compiles
*> origin: languages/cobol/tests/cobol/test_declaratives_use_sections.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
DECLARATIVES.
ERR-SEC SECTION.
    USE AFTER STANDARD ERROR PROCEDURE ON WS-FILE.
END DECLARATIVES.
MAIN-SEC SECTION.
    DISPLAY "RUN".
    STOP RUN.

