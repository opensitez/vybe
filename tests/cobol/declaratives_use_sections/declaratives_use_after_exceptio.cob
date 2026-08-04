*> vybe-test: cobol/declaratives_use_sections/declaratives_use_after_exception_compiles
*> origin: languages/cobol/tests/cobol/test_declaratives_use_sections.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
DECLARATIVES.
EX-SEC SECTION.
    USE FOR DEBUGGING ON ALL PROCEDURES.
END DECLARATIVES.
MAIN-SEC SECTION.
    DISPLAY "RUN".
    STOP RUN.

