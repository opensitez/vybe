*> vybe-test: cobol/category_procedure_division_alter_statement/alter_allows_hyphenated_paragraph_names
*> origin: languages/cobol/tests/cobol/test_category_procedure_division_alter_statement.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. AL5.
       PROCEDURE DIVISION.
           ALTER START-POINT TO PROCEED TO ALT-PATH.
           GO TO START-POINT.
       START-POINT.
           DISPLAY "START".
           STOP RUN.
       ALT-PATH.
           DISPLAY "ALT-PATH".
           STOP RUN.

