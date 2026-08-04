*> vybe-test: cobol/copy_replacing/replace_off_compiles
*> origin: languages/cobol/tests/cobol/test_copy_replacing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. CPY4.
PROCEDURE DIVISION.
    REPLACE ==A== BY ==B==.
    REPLACE OFF.
    STOP RUN.

