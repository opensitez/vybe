*> vybe-test: cobol/exceptions_error_paths/compute_size_error_branch_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 9.
01 B PIC 9 VALUE 9.
01 C PIC 9 VALUE 0.
PROCEDURE DIVISION.
    COMPUTE C = A ** B
        ON SIZE ERROR DISPLAY "SE"
    END-COMPUTE.
    STOP RUN.

