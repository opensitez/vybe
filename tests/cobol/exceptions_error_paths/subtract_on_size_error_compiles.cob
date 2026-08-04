*> vybe-test: cobol/exceptions_error_paths/subtract_on_size_error_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(2) VALUE 10.
01 B PIC 9(2) VALUE 20.
01 C PIC 9(2).
PROCEDURE DIVISION.
    SUBTRACT B FROM A GIVING C
        ON SIZE ERROR DISPLAY "SE"
    END-SUBTRACT.
    STOP RUN.

