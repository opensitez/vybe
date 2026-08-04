*> vybe-test: cobol/symbolic_characters/symbolic_characters_in_display_compiles
*> origin: languages/cobol/tests/cobol/test_symbolic_characters.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. SYM3.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    SYMBOLIC CHARACTERS S-A IS 65.
PROCEDURE DIVISION.
    DISPLAY S-A.
    STOP RUN.

