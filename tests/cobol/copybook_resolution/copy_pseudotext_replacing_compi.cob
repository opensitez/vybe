*> vybe-test: cobol/copybook_resolution/copy_pseudotext_replacing_compiles
*> origin: languages/cobol/tests/cobol/test_copybook_resolution.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CBR8.
PROCEDURE DIVISION.
    COPY CUSTOMER-REC REPLACING ==CUST-ID== BY ==ORDER-ID==.
    STOP RUN.

