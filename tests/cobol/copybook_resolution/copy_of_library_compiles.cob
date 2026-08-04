*> vybe-test: cobol/copybook_resolution/copy_of_library_compiles
*> origin: languages/cobol/tests/cobol/test_copybook_resolution.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CBR2.
PROCEDURE DIVISION.
    COPY DATE-UTILS OF COMMON-LIB.
    STOP RUN.

