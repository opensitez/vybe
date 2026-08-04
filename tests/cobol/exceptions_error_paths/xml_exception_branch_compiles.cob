*> vybe-test: cobol/exceptions_error_paths/xml_exception_branch_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(50).
01 R PIC X(5).
PROCEDURE DIVISION.
    XML GENERATE X FROM R
        ON EXCEPTION DISPLAY "XERR"
    END-XML.
    STOP RUN.

