*> vybe-test: cobol/category_object_oriented_advanced/test_oo_invoke_factory
*> origin: languages/cobol/tests/cobol/test_category_object_oriented_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. ENVIRONMENT DIVISION. CONFIGURATION SECTION. REPOSITORY. CLASS C IS 'C'. PROCEDURE DIVISION. INVOKE C 'NEW' RETURNING O. DISPLAY 'OK'. STOP RUN.

