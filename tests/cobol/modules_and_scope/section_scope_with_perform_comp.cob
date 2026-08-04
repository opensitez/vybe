*> vybe-test: cobol/modules_and_scope/section_scope_with_perform_compiles
*> origin: languages/cobol/tests/cobol/test_modules_and_scope.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    PERFORM SECTION-ONE.
    STOP RUN.
SECTION-ONE SECTION.
    DISPLAY "ONE".

