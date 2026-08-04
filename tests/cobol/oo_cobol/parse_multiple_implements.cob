*> vybe-test: cobol/oo_cobol/parse_multiple_implements
*> origin: languages/cobol/tests/cobol/test_oo_cobol.rs

IDENTIFICATION DIVISION.
CLASS-ID. MY-OBJ IMPLEMENTS PRINTABLE, COMPARABLE.
OBJECT.
METHOD-ID. TO-STRING.
PROCEDURE DIVISION.
    DISPLAY "Object".
END METHOD TO-STRING.
END OBJECT.
END CLASS MY-OBJ.

