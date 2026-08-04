*> vybe-test: cobol/copy_replacing/replace_then_copy_compiles
*> origin: languages/cobol/tests/cobol/test_copy_replacing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CPY8.
PROCEDURE DIVISION.
    REPLACE ==A== BY ==B==.
    COPY BASIC-BOOK.
    REPLACE OFF.
    STOP RUN.

