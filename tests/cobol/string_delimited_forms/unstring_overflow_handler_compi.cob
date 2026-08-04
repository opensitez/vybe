*> vybe-test: cobol/string_delimited_forms/unstring_overflow_handler_compiles
*> origin: languages/cobol/tests/cobol/test_string_delimited_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC X(10) VALUE "A,B,C,D,E".
01 F1 PIC X(5).
PROCEDURE DIVISION.
    UNSTRING SRC DELIMITED BY "," INTO F1
    ON OVERFLOW
        DISPLAY "OVERFLOW"
    END-UNSTRING.
    STOP RUN.

