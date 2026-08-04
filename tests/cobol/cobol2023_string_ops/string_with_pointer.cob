*> vybe-test: cobol/cobol2023_string_ops/string_with_pointer
*> origin: languages/cobol/tests/cobol/test_cobol2023_string_ops.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RESULT PIC X(50) VALUE SPACES.
01 WS-PTR PIC 9(3) VALUE 1.
PROCEDURE DIVISION.
    STRING "Hello" DELIMITED BY SIZE
           " "     DELIMITED BY SIZE
           "World" DELIMITED BY SIZE
        INTO WS-RESULT
        WITH POINTER WS-PTR.
    DISPLAY WS-RESULT.
    STOP RUN.

