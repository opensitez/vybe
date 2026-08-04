*> vybe-test: cobol/accept_forms/goback_in_subprogram_context_compiles
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. SUB.
PROCEDURE DIVISION.
    DISPLAY "SUB".
    GOBACK.

