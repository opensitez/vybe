*> vybe-test: cobol/cobol2023_nested/exit_method
*> origin: languages/cobol/tests/cobol/test_cobol2023_nested.rs

IDENTIFICATION DIVISION.
CLASS-ID. MY-CLASS.
OBJECT.
METHOD-ID. PROCESS.
PROCEDURE DIVISION.
    DISPLAY "Start".
    EXIT METHOD.
    DISPLAY "Never reached".
END METHOD PROCESS.
END OBJECT.
END CLASS MY-CLASS.

