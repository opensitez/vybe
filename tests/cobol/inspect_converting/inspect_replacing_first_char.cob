*> vybe-test: cobol/inspect_converting/inspect_replacing_first_char
*> origin: languages/cobol/tests/cobol/test_inspect_converting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(10) VALUE "AABABC".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    INSPECT S REPLACING FIRST "A" BY "X".
    DISPLAY S.
    MOVE SPACES TO WS-VYBE-L
    STRING S DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "XABABC    "
        DISPLAY "FAIL: want [XABABC    ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

