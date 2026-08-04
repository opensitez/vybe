*> vybe-test: cobol/category_procedure_division_alter_statement/alter_parse_with_section_like_names
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_alter_statement.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. AL6.
       PROCEDURE DIVISION.
           ALTER ROUTE-ONE TO PROCEED TO ROUTE-TWO.
       ROUTE-ONE.
           GO TO END-POINT.
       ROUTE-TWO.
           DISPLAY "ROUTE-TWO".
       END-POINT.
           STOP RUN.

