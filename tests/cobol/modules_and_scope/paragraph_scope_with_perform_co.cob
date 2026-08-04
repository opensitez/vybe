*> vybe-test: cobol/modules_and_scope/paragraph_scope_with_perform_compiles
*> origin: languages/cobol/tests/cobol/test_modules_and_scope.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    PERFORM PARA-ONE.
    STOP RUN.
PARA-ONE.
    DISPLAY "ONE".

