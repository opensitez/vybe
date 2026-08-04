*> vybe-test: cobol/exceptions_error_paths/string_overflow_branch_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(5) VALUE "ABCDE".
01 B PIC X(5) VALUE "FGHIJ".
01 O PIC X(5).
PROCEDURE DIVISION.
    STRING A DELIMITED BY SIZE B DELIMITED BY SIZE INTO O
        ON OVERFLOW DISPLAY "OV"
    END-STRING.
    STOP RUN.

