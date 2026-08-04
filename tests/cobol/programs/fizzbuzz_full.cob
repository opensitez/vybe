*> vybe-test: cobol/programs/fizzbuzz_full
*> origin: languages/cobol/tests/cobol/test_programs.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. FIZZBUZZ.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-I      PIC 9(3) VALUE 0.
01 WS-MOD3   PIC 9(3) VALUE 0.
01 WS-MOD5   PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 100
        COMPUTE WS-MOD3 = FUNCTION MOD(WS-I 3)
        COMPUTE WS-MOD5 = FUNCTION MOD(WS-I 5)
        EVALUATE TRUE
            WHEN WS-MOD3 = 0 AND WS-MOD5 = 0
                DISPLAY "FizzBuzz"
            WHEN WS-MOD3 = 0
                DISPLAY "Fizz"
            WHEN WS-MOD5 = 0
                DISPLAY "Buzz"
            WHEN OTHER
                DISPLAY WS-I
        END-EVALUATE
    END-PERFORM.
    STOP RUN.

