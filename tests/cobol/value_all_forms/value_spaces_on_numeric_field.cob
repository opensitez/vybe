*> vybe-test: cobol/value_all_forms/value_spaces_on_numeric_field
*> ⛔ INVALID COBOL. cobc: `error: invalid VALUE clause` — `VALUE SPACES` on a
*> `PIC 9(5)` is not legal, so neither our old `00000` nor our new `00NaN` is
*> a right answer; there is no right answer to a program that cannot compile.
*> Left FAILING deliberately rather than pinned to whatever we happen to
*> produce — encoding an artifact of ours as the expectation would make this
*> test permanently green and permanently meaningless.
*> origin: languages/cobol/tests/cobol/test_value_all_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(5) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY N.
    MOVE SPACES TO WS-VYBE-L
    STRING N DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "00000"
        DISPLAY "FAIL: want [00000] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

