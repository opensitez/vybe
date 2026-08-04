*> vybe-test: cobol/string_and_unstring_extended/string_on_overflow_executes_overflow_branch
*> origin: languages/cobol/tests/cobol/test_string_and_unstring_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(3) VALUE "ABC".
01 WS-B PIC X(3) VALUE "DEF".
01 WS-R PIC X(3) VALUE SPACES.
PROCEDURE DIVISION.

    STRING WS-A DELIMITED BY SIZE
           WS-B DELIMITED BY SIZE
           INTO WS-R
      ON OVERFLOW DISPLAY "OVF"
      NOT ON OVERFLOW DISPLAY "OK"
    END-STRING.
    STOP RUN.

