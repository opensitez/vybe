*> vybe-test: cobol/exceptions_error_paths/divide_size_error_branch_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 1.
01 B PIC 9 VALUE 0.
PROCEDURE DIVISION.
    DIVIDE A BY B
        ON SIZE ERROR DISPLAY "SE"
    END-DIVIDE.
    STOP RUN.

