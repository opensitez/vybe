*> vybe-test: cobol/properties_attributes_accessors/prop_getter_compiles
*> origin: languages/cobol/tests/cobol/test_properties_attributes_accessors.rs
IDENTIFICATION DIVISION.
CLASS-ID. P1.
OBJECT.
METHOD-ID. GET-A PROPERTY GET.
PROCEDURE DIVISION RETURNING R.
    MOVE 1 TO R.
END METHOD GET-A.
END OBJECT.
END CLASS P1.

