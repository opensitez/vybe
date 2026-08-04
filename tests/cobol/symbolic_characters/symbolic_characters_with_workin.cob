*> vybe-test: cobol/symbolic_characters/symbolic_characters_with_working_storage_compiles
*> origin: languages/cobol/tests/cobol/test_symbolic_characters.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. SYM8.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    SYMBOLIC CHARACTERS X-SYM IS 88.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 CH PIC X.
PROCEDURE DIVISION.
    MOVE X-SYM TO CH.
    STOP RUN.

