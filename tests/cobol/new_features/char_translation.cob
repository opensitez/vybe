*> vybe-test: cobol/new_features/char_translation
*> origin: languages/cobol/tests/cobol/test_new_features.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. CONVERT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(30) VALUE "Hello World 123".
01 WS-COPY PIC X(30).
PROCEDURE DIVISION.
    MOVE WS-TEXT TO WS-COPY.
    INSPECT WS-COPY CONVERTING "abcdefghijklmnopqrstuvwxyz"
                    TO "ABCDEFGHIJKLMNOPQRSTUVWXYZ".
    DISPLAY "Original:  " WS-TEXT.
    DISPLAY "Converted: " WS-COPY.
    STOP RUN.

