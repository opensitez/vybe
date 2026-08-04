*> vybe-test: cobol/properties_attributes_accessors/prop_factory_and_property_compiles
*> origin: languages/cobol/tests/cobol/test_properties_attributes_accessors.rs
IDENTIFICATION DIVISION.
CLASS-ID. P9.
FACTORY.
METHOD-ID. NEWP.
PROCEDURE DIVISION RETURNING O.
    DISPLAY "N".
END METHOD NEWP.
END FACTORY.
OBJECT.
METHOD-ID. GET-V PROPERTY GET.
PROCEDURE DIVISION RETURNING R.
    MOVE 1 TO R.
END METHOD GET-V.
END OBJECT.
END CLASS P9.

