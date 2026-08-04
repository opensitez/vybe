*> vybe-test: cobol/exceptions_error_paths/unstring_overflow_branch_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(3) VALUE "A,B".
01 A PIC X(1).
PROCEDURE DIVISION.
    UNSTRING S DELIMITED BY "," INTO A
        ON OVERFLOW DISPLAY "OV"
    END-UNSTRING.
    STOP RUN.

