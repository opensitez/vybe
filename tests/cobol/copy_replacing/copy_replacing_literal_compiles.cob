*> vybe-test: cobol/copy_replacing/copy_replacing_literal_compiles
*> origin: languages/cobol/tests/cobol/test_copy_replacing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CPY6.
PROCEDURE DIVISION.
    COPY TEXT-BOOK REPLACING "OLD" BY "NEW".
    STOP RUN.

