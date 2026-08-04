*> vybe-test: cobol/copybook_resolution/copy_in_alt_library_compiles
*> origin: languages/cobol/tests/cobol/test_copybook_resolution.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CBR6.
PROCEDURE DIVISION.
    COPY ITEM-REC IN DATA-LIB.
    STOP RUN.

