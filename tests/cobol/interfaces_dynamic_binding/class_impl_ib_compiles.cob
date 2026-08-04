*> vybe-test: cobol/interfaces_dynamic_binding/class_impl_ib_compiles
*> origin: languages/cobol/tests/cobol/test_interfaces_dynamic_binding.rs
IDENTIFICATION DIVISION.
CLASS-ID. CB IMPLEMENTS IB.
OBJECT.
METHOD-ID. M2.
PROCEDURE DIVISION.
    DISPLAY "2".
END METHOD M2.
END OBJECT.
END CLASS CB.

