*> vybe-test: cobol/runtime_environment/display_upon_environment_value_compiles
*> origin: languages/cobol/tests/cobol/test_runtime_environment.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. ENV3.
PROCEDURE DIVISION.
    DISPLAY "X" UPON ENVIRONMENT-VALUE.
    STOP RUN.

