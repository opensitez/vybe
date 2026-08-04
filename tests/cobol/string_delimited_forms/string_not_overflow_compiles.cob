*> vybe-test: cobol/string_delimited_forms/string_not_overflow_compiles
*> origin: languages/cobol/tests/cobol/test_string_delimited_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(5) VALUE "HELLO".
01 R PIC X(10) VALUE SPACES.
PROCEDURE DIVISION.
    STRING A DELIMITED BY SIZE INTO R
    NOT ON OVERFLOW
        DISPLAY "OK"
    END-STRING.
    STOP RUN.

