*> vybe-test: cobol/category_data_division_renames/test_renames_parse_alias_from_redefined_area
*> origin: languages/cobol/tests/cobol/test_category_data_division_renames.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R. 05 A PIC X VALUE 'A'. 05 B PIC X VALUE 'B'. 05 C PIC X VALUE 'C'. 66 ALIAS RENAMES A THRU B. 01 R2 REDEFINES R PIC XX. 66 ALIAS2 RENAMES R2. PROCEDURE DIVISION. MOVE 'ZZ' TO ALIAS2 DISPLAY ALIAS.
    MOVE SPACES TO WS-VYBE-L
    STRING ALIAS DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ZZ"
        DISPLAY "FAIL: want [ZZ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

