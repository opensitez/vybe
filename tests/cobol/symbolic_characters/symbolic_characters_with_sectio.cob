*> vybe-test: cobol/symbolic_characters/symbolic_characters_with_sectioned_program_compiles
*> origin: languages/cobol/tests/cobol/test_symbolic_characters.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. SYM10.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    SYMBOLIC CHARACTERS NL IS 10.
PROCEDURE DIVISION.
MAIN SECTION.
    DISPLAY NL.
    STOP RUN.

