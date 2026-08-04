*> vybe-test: cobol/string_delimited_forms/string_pointer_updated_compiles
*> origin: languages/cobol/tests/cobol/test_string_delimited_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(3) VALUE "ABC".
01 B PIC X(3) VALUE "DEF".
01 R PIC X(20) VALUE SPACES.
01 PTR PIC 9(3) VALUE 1.
PROCEDURE DIVISION.
    STRING A DELIMITED BY SIZE INTO R WITH POINTER PTR.
    STRING B DELIMITED BY SIZE INTO R WITH POINTER PTR.
    STOP RUN.

