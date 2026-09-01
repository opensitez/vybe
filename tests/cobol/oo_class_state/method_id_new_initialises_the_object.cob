*> vybe-test: cobol/oo_class_state/method_id_new_initialises_the_object
*>
*> `METHOD-ID. NEW` is the constructor — `constructor_name = "NEW"` in the
*> COBOL profile says so, and `normalize_class.rs` routes it through
*> `push_constructor`. The LEGACY normalizer COBOL was silently taking had no
*> constructor concept at all, so this body ran as an ordinary method that
*> nobody called and the VALUE clause was the only initialiser.
IDENTIFICATION DIVISION.
CLASS-ID. COUNTER.
OBJECT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COUNT PIC 9(4) VALUE 0.
METHOD-ID. NEW.
PROCEDURE DIVISION.
    MOVE 42 TO WS-COUNT.
END METHOD NEW.
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
    INVOKE O GET-COUNT RETURNING R.
    IF R NOT = 42
        DISPLAY "FAIL: constructor ran want [42] got [" R "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    DISPLAY "OK".
    STOP RUN.
