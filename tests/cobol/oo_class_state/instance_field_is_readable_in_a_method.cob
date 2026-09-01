*> vybe-test: cobol/oo_class_state/instance_field_is_readable_in_a_method
*>
*> A bare class-level data name inside a METHOD is that object's FIELD.
*> Every other OO test in this tree asserts only that the program exits 0, so
*> this whole layer could be — and was — inert: `WS-COUNT` compiled to a
*> `global.get`, the constructor wrote the instance field and nothing ever
*> read it, and all 174 of them still passed. This one reads a value back.
IDENTIFICATION DIVISION.
CLASS-ID. COUNTER.
OBJECT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COUNT PIC 9(4) VALUE 7.
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
    IF R NOT = 7
        DISPLAY "FAIL: field read want [7] got [" R "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    DISPLAY "OK".
    STOP RUN.
