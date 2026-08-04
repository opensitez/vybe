*> vybe-test: cobol/oo_properties_and_attributes/class_write_only_property_compiles
*> origin: languages/cobol/tests/cobol/test_oo_properties_and_attributes.rs

IDENTIFICATION DIVISION.
CLASS-ID. WRITE-ONLY.
OBJECT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NOTE PIC X(30).
METHOD-ID. SET-NOTE PROPERTY SET.
PROCEDURE DIVISION USING WS-IN.
    MOVE WS-IN TO WS-NOTE.
END METHOD SET-NOTE.
END OBJECT.
END CLASS WRITE-ONLY.

