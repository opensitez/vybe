*> vybe-test: cobol/accept_forms/exit_in_inline_loop_compiles
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3
        IF I = 2
            CONTINUE
        END-IF
    END-PERFORM.
    STOP RUN.

