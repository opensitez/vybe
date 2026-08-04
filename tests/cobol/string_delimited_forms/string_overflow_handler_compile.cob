*> vybe-test: cobol/string_delimited_forms/string_overflow_handler_compiles
*> origin: languages/cobol/tests/cobol/test_string_delimited_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(20) VALUE "LONG VALUE HERE      ".
01 R PIC X(5) VALUE SPACES.
PROCEDURE DIVISION.
    STRING A DELIMITED BY SIZE INTO R
    ON OVERFLOW
        DISPLAY "OVERFLOW"
    END-STRING.
    STOP RUN.

