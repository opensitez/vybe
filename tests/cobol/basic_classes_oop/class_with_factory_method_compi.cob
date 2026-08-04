*> vybe-test: cobol/basic_classes_oop/class_with_factory_method_compiles
*> origin: languages/cobol/tests/cobol/test_basic_classes_oop.rs

IDENTIFICATION DIVISION.
CLASS-ID. COUNTER-FACTORY.
FACTORY.
METHOD-ID. CREATE-COUNTER.
PROCEDURE DIVISION RETURNING WS-INSTANCE.
    DISPLAY "CREATE".
END METHOD CREATE-COUNTER.
END FACTORY.
END CLASS COUNTER-FACTORY.

