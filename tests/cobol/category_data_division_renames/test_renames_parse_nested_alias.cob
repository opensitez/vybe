*> vybe-test: cobol/category_data_division_renames/test_renames_parse_nested_alias_chain
*> origin: languages/cobol/tests/cobol/test_category_data_division_renames.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R. 05 OUT. 10 A PIC X VALUE 'A'. 10 B PIC X VALUE 'B'. 10 C PIC X VALUE 'C'. 66 ALIAS RENAMES A THRU C. 66 ALIAS2 RENAMES A THRU B. PROCEDURE DIVISION. MOVE 'WX' TO ALIAS2 DISPLAY ALIAS.
    MOVE SPACES TO WS-VYBE-L
    STRING ALIAS DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "WXC"
        DISPLAY "FAIL: want [WXC] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

