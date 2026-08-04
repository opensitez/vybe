*> vybe-test: cobol/symbolic_characters/symbolic_characters_multiple_compiles
*> origin: languages/cobol/tests/cobol/test_symbolic_characters.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. SYM2.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    SYMBOLIC CHARACTERS C1 IS 1 C2 IS 2 C3 IS 3.
PROCEDURE DIVISION.
    STOP RUN.

