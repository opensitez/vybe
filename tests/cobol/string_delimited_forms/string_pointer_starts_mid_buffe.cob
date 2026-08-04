*> vybe-test: cobol/string_delimited_forms/string_pointer_starts_mid_buffer
*> origin: languages/cobol/tests/cobol/test_string_delimited_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(5) VALUE "HELLO".
01 R PIC X(20) VALUE SPACES.
01 PTR PIC 9(3) VALUE 6.
PROCEDURE DIVISION.
    STRING A DELIMITED BY SIZE INTO R WITH POINTER PTR.
    STOP RUN.

