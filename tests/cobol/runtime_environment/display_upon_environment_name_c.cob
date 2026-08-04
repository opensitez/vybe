*> vybe-test: cobol/runtime_environment/display_upon_environment_name_compiles
*> origin: languages/cobol/tests/cobol/test_runtime_environment.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. ENV2.
PROCEDURE DIVISION.
    DISPLAY "X" UPON ENVIRONMENT-NAME.
    STOP RUN.

