*> vybe-test: cobol/oo_class_state/a_method_mutates_its_own_objects_field
*>
*> A method WRITES the instance field, and a later method READS the written
*> value. The write and the read have to land on the same storage; while a
*> bare class data name resolved to a module global they landed on the same
*> WRONG storage, which is indistinguishable from working until two objects
*> exist — see `two_objects_do_not_share_state`.
IDENTIFICATION DIVISION.
CLASS-ID. COUNTER.
OBJECT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COUNT PIC 9(4) VALUE 0.
METHOD-ID. BUMP.
PROCEDURE DIVISION.
    ADD 3 TO WS-COUNT.
END METHOD BUMP.
METHOD-ID. GET-COUNT.
PROCEDURE DIVISION RETURNING WS-RESULT.
    MOVE WS-COUNT TO WS-RESULT.
END METHOD GET-COUNT.
END OBJECT.
END CLASS COUNTER.
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 O USAGE OBJECT REFERENCE COUNTER.
01 R PIC 9(4).
PROCEDURE DIVISION.
    INVOKE COUNTER NEW RETURNING O.
    INVOKE O BUMP.
    INVOKE O BUMP.
    INVOKE O GET-COUNT RETURNING R.
    IF R NOT = 6
        DISPLAY "FAIL: two bumps want [6] got [" R "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    DISPLAY "OK".
    STOP RUN.
