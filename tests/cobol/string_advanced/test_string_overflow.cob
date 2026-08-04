*> vybe-test: cobol/string_advanced/test_string_overflow
*> origin: languages/cobol/tests/cobol/test_string_advanced.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(10) VALUE "ABCDEFGHIJ".
01 WS-B PIC X(10) VALUE "KLQRSTUVWX".
01 WS-DST PIC X(15).
PROCEDURE DIVISION.

    STRING WS-A WS-B DELIMITED BY SIZE INTO WS-DST
        ON OVERFLOW
            DISPLAY "OVERFLOW"
        NOT ON OVERFLOW
            DISPLAY "OK"
    END-STRING.
    STOP RUN.

