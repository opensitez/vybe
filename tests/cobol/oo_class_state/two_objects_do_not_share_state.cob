*> vybe-test: cobol/oo_class_state/two_objects_do_not_share_state
*>
*> The discriminator the other OO tests cannot be: if a class field is really
*> a module global, ONE storage is shared by every instance and this reads 5
*> for both. Instance state is what makes a class a class, so this is the
*> test that says whether the OO layer is real.
IDENTIFICATION DIVISION.
CLASS-ID. BOX.
OBJECT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-N PIC 9(4) VALUE 0.
METHOD-ID. BUMP.
PROCEDURE DIVISION.
    ADD 1 TO WS-N.
END METHOD BUMP.
METHOD-ID. GET-N.
PROCEDURE DIVISION RETURNING WS-RESULT.
    MOVE WS-N TO WS-RESULT.
END METHOD GET-N.
END OBJECT.
END CLASS BOX.
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A USAGE OBJECT REFERENCE BOX.
01 B USAGE OBJECT REFERENCE BOX.
01 RA PIC 9(4).
01 RB PIC 9(4).
PROCEDURE DIVISION.
    INVOKE BOX NEW RETURNING A.
    INVOKE BOX NEW RETURNING B.
    INVOKE A BUMP.
    INVOKE A BUMP.
    INVOKE A BUMP.
    INVOKE B BUMP.
    INVOKE A GET-N RETURNING RA.
    INVOKE B GET-N RETURNING RB.
    IF RA NOT = 3
        DISPLAY "FAIL: object A want [3] got [" RA "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    IF RB NOT = 1
        DISPLAY "FAIL: object B want [1] got [" RB "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    DISPLAY "OK".
    STOP RUN.
