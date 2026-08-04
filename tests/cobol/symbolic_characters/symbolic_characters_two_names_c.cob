*> vybe-test: cobol/symbolic_characters/symbolic_characters_two_names_compiles
*> origin: languages/cobol/tests/cobol/test_symbolic_characters.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. SYM5.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    SYMBOLIC CHARACTERS S1 IS 10 S2 IS 11.
PROCEDURE DIVISION.
    STOP RUN.

