*> vybe-test: cobol/copy_replacing/copy_replacing_partial_words_compiles
*> origin: languages/cobol/tests/cobol/test_copy_replacing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CPY7.
PROCEDURE DIVISION.
    COPY WORD-BOOK REPLACING ==CUST== BY ==ORD==.
    STOP RUN.

