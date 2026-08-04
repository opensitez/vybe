*> vybe-test: cobol/properties_attributes_accessors/prop_class_multiple_methods_compiles
*> origin: languages/cobol/tests/cobol/test_properties_attributes_accessors.rs
IDENTIFICATION DIVISION.
CLASS-ID. P14.
OBJECT.
METHOD-ID. A.
PROCEDURE DIVISION.
    DISPLAY "A".
END METHOD A.
METHOD-ID. B.
PROCEDURE DIVISION.
    DISPLAY "B".
END METHOD B.
END OBJECT.
END CLASS P14.

