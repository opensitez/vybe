*> vybe-test: cobol/declaratives_use_sections/declaratives_section_labels_compiles
*> origin: languages/cobol/tests/cobol/test_declaratives_use_sections.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
DECLARATIVES.
D-A SECTION.
    USE AFTER STANDARD ERROR PROCEDURE ON WS-FILE.
D-B SECTION.
    USE AFTER STANDARD ERROR PROCEDURE ON WS-FILE.
END DECLARATIVES.
MAIN SECTION.
    DISPLAY "OK".
    STOP RUN.

