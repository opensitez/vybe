*> vybe-test: cobol/symbolic_characters/symbolic_characters_four_names_compiles
*> origin: languages/cobol/tests/cobol/test_symbolic_characters.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. SYM6.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    SYMBOLIC CHARACTERS S1 IS 1 S2 IS 2 S3 IS 3 S4 IS 4.
PROCEDURE DIVISION.
    STOP RUN.

