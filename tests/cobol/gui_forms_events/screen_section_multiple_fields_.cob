*> vybe-test: cobol/gui_forms_events/screen_section_multiple_fields_compiles
*> origin: languages/cobol/tests/cobol/test_gui_forms_events.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. S13.
DATA DIVISION.
SCREEN SECTION.
01 SCR.
   05 LINE 1 COLUMN 1 PIC X(20) USING N1.
   05 LINE 2 COLUMN 1 PIC X(20) USING N2.
WORKING-STORAGE SECTION.
01 N1 PIC X(20).
01 N2 PIC X(20).
PROCEDURE DIVISION.
    DISPLAY SCR.
    ACCEPT SCR.
    STOP RUN.

