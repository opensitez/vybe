*> vybe-test: cobol/perform_inline_declarative/perform_varying_with_float_step_compiles
*> origin: languages/cobol/tests/cobol/test_perform_inline_declarative.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9(3)V9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING I FROM 0 BY 0.5 UNTIL I > 2
        CONTINUE
    END-PERFORM.
    STOP RUN.

