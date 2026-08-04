*> vybe-test: cobol/copybook_resolution/copy_in_library_compiles
*> origin: languages/cobol/tests/cobol/test_copybook_resolution.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CBR1.
PROCEDURE DIVISION.
    COPY CUSTOMER-REC IN COMMON-LIB.
    STOP RUN.

