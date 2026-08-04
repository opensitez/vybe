*> vybe-test: cobol/category_object_oriented_advanced/test_oo_exception
*> origin: languages/cobol/tests/cobol/test_category_object_oriented_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. INVOKE O 'M1' ON EXCEPTION DISPLAY 'E' NOT ON EXCEPTION DISPLAY 'N' END-INVOKE. STOP RUN.

