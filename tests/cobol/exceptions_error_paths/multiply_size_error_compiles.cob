*> vybe-test: cobol/exceptions_error_paths/multiply_size_error_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(4) VALUE 200.
01 B PIC 9(4) VALUE 300.
01 C PIC 9(4).
PROCEDURE DIVISION.
    MULTIPLY A BY B GIVING C
        ON SIZE ERROR DISPLAY "MUL-SE"
    END-MULTIPLY.
    STOP RUN.

