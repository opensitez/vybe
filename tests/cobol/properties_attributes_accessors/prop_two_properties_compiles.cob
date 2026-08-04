*> vybe-test: cobol/properties_attributes_accessors/prop_two_properties_compiles
*> origin: languages/cobol/tests/cobol/test_properties_attributes_accessors.rs
IDENTIFICATION DIVISION.
CLASS-ID. P6.
OBJECT.
METHOD-ID. GET-A PROPERTY GET.
PROCEDURE DIVISION RETURNING R.
    MOVE 1 TO R.
END METHOD GET-A.
METHOD-ID. GET-B PROPERTY GET.
PROCEDURE DIVISION RETURNING R.
    MOVE 2 TO R.
END METHOD GET-B.
END OBJECT.
END CLASS P6.

