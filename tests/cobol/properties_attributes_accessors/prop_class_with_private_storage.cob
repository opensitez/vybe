*> vybe-test: cobol/properties_attributes_accessors/prop_class_with_private_storage_compiles
*> origin: languages/cobol/tests/cobol/test_properties_attributes_accessors.rs
IDENTIFICATION DIVISION.
CLASS-ID. P11.
OBJECT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 V PIC X(10).
METHOD-ID. SET-V PROPERTY SET.
PROCEDURE DIVISION USING I.
    MOVE I TO V.
END METHOD SET-V.
END OBJECT.
END CLASS P11.

