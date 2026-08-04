*> vybe-test: cobol/symbolic_characters/symbolic_characters_with_display_usage_compiles
*> origin: languages/cobol/tests/cobol/test_symbolic_characters.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. SYM7.
ENVIRONMENT DIVISION.
CONFIGURATION SECTION.
SPECIAL-NAMES.
    SYMBOLIC CHARACTERS TAB-SYM IS 9.
PROCEDURE DIVISION.
    DISPLAY TAB-SYM.
    STOP RUN.

