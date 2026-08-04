*> vybe-test: cobol/properties_attributes_accessors/prop_numeric_property_compiles
*> origin: languages/cobol/tests/cobol/test_properties_attributes_accessors.rs
IDENTIFICATION DIVISION.
CLASS-ID. P5.
OBJECT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 V PIC 9(5).
METHOD-ID. SET-V PROPERTY SET.
PROCEDURE DIVISION USING I.
    MOVE I TO V.
END METHOD SET-V.
END OBJECT.
END CLASS P5.

