*> vybe-test: cobol/events_and_callbacks_expanded/nested_perform_blocks_compiles
*> origin: languages/cobol/tests/cobol/test_events_and_callbacks_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 0.
01 J PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 2
        PERFORM VARYING J FROM 1 BY 1 UNTIL J > 2
            DISPLAY I
        END-PERFORM
    END-PERFORM.
    STOP RUN.

