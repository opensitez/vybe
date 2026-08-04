*> vybe-test: cobol/exceptions_error_paths/add_on_size_error_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 999.
01 B PIC 9(3) VALUE 2.
01 C PIC 9(3).
PROCEDURE DIVISION.
    ADD A TO B GIVING C
        ON SIZE ERROR DISPLAY "SE"
        NOT ON SIZE ERROR DISPLAY "OK"
    END-ADD.
    STOP RUN.

