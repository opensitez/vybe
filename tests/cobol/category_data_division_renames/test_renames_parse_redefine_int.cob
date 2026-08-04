*> vybe-test: cobol/category_data_division_renames/test_renames_parse_redefine_interaction
*> origin: languages/cobol/tests/cobol/test_category_data_division_renames.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R. 05 A PIC X VALUE 'X'. 05 B PIC X VALUE 'Y'. 66 ALIAS RENAMES A THRU B. 01 DUP REDEFINES R PIC X(2). 66 ALIAS2 RENAMES DUP. PROCEDURE DIVISION. MOVE 'ZZ' TO ALIAS DISPLAY ALIAS2.
    MOVE SPACES TO WS-VYBE-L
    STRING ALIAS2 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ZZ"
        DISPLAY "FAIL: want [ZZ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

