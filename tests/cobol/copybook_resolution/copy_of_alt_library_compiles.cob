*> vybe-test: cobol/copybook_resolution/copy_of_alt_library_compiles
*> origin: languages/cobol/tests/cobol/test_copybook_resolution.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CBR7.
PROCEDURE DIVISION.
    COPY ITEM-REC OF DATA-LIB.
    STOP RUN.

