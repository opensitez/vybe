*> vybe-test: cobol/symbolic_characters/symbolic_characters_with_multiple_display_compiles
*> origin: languages/cobol/tests/cobol/test_symbolic_characters.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. SYM9.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    SYMBOLIC CHARACTERS C1 IS 1 C2 IS 2.
PROCEDURE DIVISION.
    DISPLAY C1 C2.
    STOP RUN.

