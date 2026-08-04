*> vybe-test: cobol/cobol2023_string_ops/string_overflow_handler
*> origin: languages/cobol/tests/cobol/test_cobol2023_string_ops.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RESULT PIC X(5) VALUE SPACES.
01 WS-LONG PIC X(20) VALUE "This is too long".
PROCEDURE DIVISION.
    STRING WS-LONG DELIMITED BY SIZE
        INTO WS-RESULT
        ON OVERFLOW DISPLAY "Overflow!"
        NOT ON OVERFLOW DISPLAY "OK"
    END-STRING.
    STOP RUN.

