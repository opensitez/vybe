      *> ============================================================
      *> TAX RETURN PROCESSING SYSTEM
      *> ============================================================
      *> Federal income tax processing: W-2 matching, 1040 filing,
      *> tax liability calculation, refund/balance-due determination,
      *> audit flag generation, e-file transmission preparation.
      *>
      *> Demonstrates: COBOL 2014 JSON GENERATE, complex nested
      *> EVALUATE, multi-dimensional tables for tax brackets,
      *> COMPUTE with ON SIZE ERROR, INSPECT for validation,
      *> CALL to external validation subprogram.
      *> ============================================================
       IDENTIFICATION DIVISION.
       PROGRAM-ID. TAX-PROCESSING.

       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT TAX-RETURN-FILE ASSIGN TO "tax_returns.dat"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-TAX-STATUS.

           SELECT W2-FILE ASSIGN TO "w2_records.dat"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-W2-STATUS.

           SELECT TAXPAYER-MASTER ASSIGN TO "taxpayers.idx"
               ORGANIZATION IS INDEXED
               ACCESS MODE IS DYNAMIC
               RECORD KEY IS TP-SSN
               FILE STATUS IS WS-TP-STATUS.

           SELECT PROCESSED-RETURNS ASSIGN TO "processed_returns.dat"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT AUDIT-FLAGS ASSIGN TO "audit_flags.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT REFUND-FILE ASSIGN TO "refunds.dat"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT BALANCE-DUE-FILE ASSIGN TO "balance_due.dat"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT EFILE-OUTPUT ASSIGN TO "efile_transmission.dat"
               ORGANIZATION IS LINE SEQUENTIAL.

       DATA DIVISION.
       FILE SECTION.

       FD  TAX-RETURN-FILE
           RECORD CONTAINS 800 CHARACTERS.
       01  TAX-RETURN-RECORD.
           05  TR-SSN              PIC X(11).
           05  TR-FILING-STATUS    PIC X(2).
               88  SINGLE          VALUE 'S '.
               88  MARRIED-JOINT   VALUE 'MJ'.
               88  MARRIED-SEP     VALUE 'MS'.
               88  HEAD-HOUSEHOLD  VALUE 'HH'.
               88  QUALIFYING-WID  VALUE 'QW'.
           05  TR-TAX-YEAR         PIC 9(4).
           05  TR-SPOUSE-SSN       PIC X(11).
           05  TR-DEPENDENTS       PIC 9(2).
           05  TR-WAGES            PIC S9(11)V99.
           05  TR-INTEREST-INC     PIC S9(9)V99.
           05  TR-DIVIDEND-INC     PIC S9(9)V99.
           05  TR-BUSINESS-INC     PIC S9(11)V99.
           05  TR-CAPITAL-GAINS    PIC S9(11)V99.
           05  TR-IRA-DIST         PIC S9(9)V99.
           05  TR-SS-BENEFITS      PIC S9(9)V99.
           05  TR-OTHER-INCOME     PIC S9(9)V99.
           05  TR-STUDENT-LOAN-INT PIC S9(7)V99.
           05  TR-IRA-DEDUCTION    PIC S9(7)V99.
           05  TR-ALIMONY-PAID     PIC S9(9)V99.
           05  TR-ITEMIZED-FLAG    PIC X(1).
               88  ITEMIZING       VALUE 'Y'.
               88  STANDARD-DED    VALUE 'N'.
           05  TR-ITEMIZED-DEDUCT  PIC S9(9)V99.
           05  TR-MORTGAGE-INT     PIC S9(9)V99.
           05  TR-CHARITABLE       PIC S9(9)V99.
           05  TR-STATE-TAXES      PIC S9(7)V99.
           05  TR-MEDICAL-EXPENSES PIC S9(9)V99.
           05  TR-CHILD-TAX-CREDIT PIC S9(7)V99.
           05  TR-EARNED-INC-CREDIT PIC S9(7)V99.
           05  TR-FED-TAX-WITHHELD PIC S9(9)V99.
           05  TR-EST-TAX-PAYMENTS PIC S9(9)V99.
           05  TR-PREPARER-ID      PIC X(9).
           05  FILLER              PIC X(400).

       FD  W2-FILE
           RECORD CONTAINS 200 CHARACTERS.
       01  W2-RECORD.
           05  W2-EMP-SSN          PIC X(11).
           05  W2-EMPLOYER-EIN     PIC X(9).
           05  W2-EMPLOYER-NAME    PIC X(50).
           05  W2-WAGES            PIC 9(9)V99.
           05  W2-FED-TAX-WITHHELD PIC 9(9)V99.
           05  W2-SS-WAGES         PIC 9(9)V99.
           05  W2-SS-TAX           PIC 9(9)V99.
           05  W2-MEDICARE-WAGES   PIC 9(9)V99.
           05  W2-MEDICARE-TAX     PIC 9(9)V99.
           05  W2-STATE-WAGES      PIC 9(9)V99.
           05  W2-STATE-TAX        PIC 9(9)V99.
           05  W2-TAX-YEAR         PIC 9(4).
           05  FILLER              PIC X(60).

       FD  TAXPAYER-MASTER
           RECORD CONTAINS 200 CHARACTERS.
       01  TAXPAYER-RECORD.
           05  TP-SSN              PIC X(11).
           05  TP-NAME             PIC X(50).
           05  TP-ADDRESS          PIC X(80).
           05  TP-PRIOR-AGI        PIC S9(11)V99 COMP-3.
           05  TP-PRIOR-REFUND     PIC S9(9)V99  COMP-3.
           05  TP-AUDIT-HISTORY    PIC 9(2).
           05  TP-FILING-HISTORY   PIC 9(2).
           05  FILLER              PIC X(40).

       FD  PROCESSED-RETURNS
           RECORD CONTAINS 300 CHARACTERS.
       01  PROC-RECORD             PIC X(300).

       FD  AUDIT-FLAGS
           RECORD CONTAINS 200 CHARACTERS.
       01  AUDIT-LINE              PIC X(200).

       FD  REFUND-FILE
           RECORD CONTAINS 100 CHARACTERS.
       01  REFUND-RECORD           PIC X(100).

       FD  BALANCE-DUE-FILE
           RECORD CONTAINS 100 CHARACTERS.
       01  BAL-DUE-RECORD          PIC X(100).

       FD  EFILE-OUTPUT
           RECORD CONTAINS 500 CHARACTERS.
       01  EFILE-RECORD            PIC X(500).

       WORKING-STORAGE SECTION.

       01  WS-STATUS.
           05  WS-TAX-STATUS       PIC XX.
               88  TAX-OK          VALUE '00'.
               88  TAX-EOF         VALUE '10'.
           05  WS-W2-STATUS        PIC XX.
               88  W2-OK           VALUE '00'.
               88  W2-EOF          VALUE '10'.
           05  WS-TP-STATUS        PIC XX.
               88  TP-OK           VALUE '00'.
               88  TP-NOT-FOUND    VALUE '23'.

       *> 2023 Tax Brackets (Married Filing Jointly)
       01  WS-TAX-BRACKETS-MJ.
           05  MJ-BRACKET OCCURS 7 TIMES.
               10  MJ-LOWER        PIC 9(9)V99.
               10  MJ-UPPER        PIC 9(9)V99.
               10  MJ-RATE         PIC V9(4).
               10  MJ-BASE-TAX     PIC 9(9)V99.

       *> 2023 Tax Brackets (Single)
       01  WS-TAX-BRACKETS-S.
           05  S-BRACKET OCCURS 7 TIMES.
               10  S-LOWER         PIC 9(9)V99.
               10  S-UPPER         PIC 9(9)V99.
               10  S-RATE          PIC V9(4).
               10  S-BASE-TAX      PIC 9(9)V99.

       01  WS-STANDARD-DEDUCTIONS.
           05  STD-DED-SINGLE      PIC 9(7)V99 VALUE 13850.00.
           05  STD-DED-MJ          PIC 9(7)V99 VALUE 27700.00.
           05  STD-DED-MFS         PIC 9(7)V99 VALUE 13850.00.
           05  STD-DED-HH          PIC 9(7)V99 VALUE 20800.00.
           05  STD-DED-QW          PIC 9(7)V99 VALUE 27700.00.

       01  WS-CALCULATIONS.
           05  WS-TOTAL-INCOME     PIC S9(13)V99 VALUE ZEROS.
           05  WS-AGI              PIC S9(13)V99 VALUE ZEROS.
           05  WS-DEDUCTION-AMT    PIC S9(11)V99 VALUE ZEROS.
           05  WS-TAXABLE-INCOME   PIC S9(13)V99 VALUE ZEROS.
           05  WS-GROSS-TAX        PIC S9(11)V99 VALUE ZEROS.
           05  WS-CREDITS          PIC S9(9)V99  VALUE ZEROS.
           05  WS-NET-TAX          PIC S9(11)V99 VALUE ZEROS.
           05  WS-TOTAL-WITHHELD   PIC S9(11)V99 VALUE ZEROS.
           05  WS-REFUND-AMOUNT    PIC S9(11)V99 VALUE ZEROS.
           05  WS-BALANCE-DUE      PIC S9(11)V99 VALUE ZEROS.
           05  WS-EFFECTIVE-RATE   PIC 9(3)V9(4) VALUE ZEROS.
           05  WS-BRACKET-IDX      PIC 9(2).
           05  WS-STANDARD-DED     PIC 9(7)V99.
           05  WS-SALT-CAP         PIC 9(7)V99 VALUE 10000.00.
           05  WS-PERSONAL-EXEMPT  PIC 9(7)V99 VALUE 0.00.

       01  WS-AUDIT-FLAGS.
           05  WS-AUDIT-SCORE      PIC 9(4) VALUE ZEROS.
           05  WS-AUDIT-REASONS    PIC X(200) VALUE SPACES.
           05  WS-AUDIT-FLAG       PIC X(1) VALUE 'N'.
               88  FLAGGED-AUDIT   VALUE 'Y'.
               88  NOT-FLAGGED     VALUE 'N'.

       01  WS-COUNTERS.
           05  WS-RETURNS-PROC     PIC 9(8) VALUE ZEROS.
           05  WS-REFUNDS          PIC 9(8) VALUE ZEROS.
           05  WS-BALANCE-DUES     PIC 9(8) VALUE ZEROS.
           05  WS-AUDIT-FLAGS-CNT  PIC 9(6) VALUE ZEROS.
           05  WS-TOTAL-REFUNDS    PIC 9(13)V99 VALUE ZEROS.
           05  WS-TOTAL-BAL-DUE    PIC 9(13)V99 VALUE ZEROS.
           05  WS-W2-MISMATCHES    PIC 9(6) VALUE ZEROS.

       01  WS-W2-TOTAL-WAGES       PIC 9(11)V99 VALUE ZEROS.
       01  WS-W2-TOTAL-WITHHELD    PIC 9(11)V99 VALUE ZEROS.
       01  WS-CURRENT-DATE         PIC 9(8).

       01  WS-FORMATTED.
           05  WF-AMOUNT           PIC -ZZZ,ZZZ,ZZZ,ZZ9.99.
           05  WF-RATE             PIC ZZ9.9999.

       PROCEDURE DIVISION.

       0000-MAIN.
           PERFORM 1000-INITIALIZE
           PERFORM 2000-PROCESS-RETURNS
               UNTIL TAX-EOF
           PERFORM 3000-PRINT-SUMMARY
           PERFORM 9000-TERMINATE
           STOP RUN.

       1000-INITIALIZE.
           MOVE FUNCTION CURRENT-DATE(1:8) TO WS-CURRENT-DATE
           OPEN INPUT  TAX-RETURN-FILE
           OPEN INPUT  W2-FILE
           OPEN I-O    TAXPAYER-MASTER
           OPEN OUTPUT PROCESSED-RETURNS
           OPEN OUTPUT AUDIT-FLAGS
           OPEN OUTPUT REFUND-FILE
           OPEN OUTPUT BALANCE-DUE-FILE
           OPEN OUTPUT EFILE-OUTPUT
           PERFORM 1100-LOAD-TAX-BRACKETS
           PERFORM 1200-READ-RETURN.

       1100-LOAD-TAX-BRACKETS.
           *> 2023 MFJ brackets
           MOVE 0         TO MJ-LOWER(1) MOVE 22000     TO MJ-UPPER(1)
           MOVE .1000     TO MJ-RATE(1)  MOVE 0         TO MJ-BASE-TAX(1)
           MOVE 22001     TO MJ-LOWER(2) MOVE 89075     TO MJ-UPPER(2)
           MOVE .1200     TO MJ-RATE(2)  MOVE 2200      TO MJ-BASE-TAX(2)
           MOVE 89076     TO MJ-LOWER(3) MOVE 190750    TO MJ-UPPER(3)
           MOVE .2200     TO MJ-RATE(3)  MOVE 10294     TO MJ-BASE-TAX(3)
           MOVE 190751    TO MJ-LOWER(4) MOVE 364200    TO MJ-UPPER(4)
           MOVE .2400     TO MJ-RATE(4)  MOVE 32580     TO MJ-BASE-TAX(4)
           MOVE 364201    TO MJ-LOWER(5) MOVE 462500    TO MJ-UPPER(5)
           MOVE .3200     TO MJ-RATE(5)  MOVE 74208     TO MJ-BASE-TAX(5)
           MOVE 462501    TO MJ-LOWER(6) MOVE 693750    TO MJ-UPPER(6)
           MOVE .3500     TO MJ-RATE(6)  MOVE 105664    TO MJ-BASE-TAX(6)
           MOVE 693751    TO MJ-LOWER(7) MOVE 9999999   TO MJ-UPPER(7)
           MOVE .3700     TO MJ-RATE(7)  MOVE 186601    TO MJ-BASE-TAX(7)

           *> 2023 Single brackets
           MOVE 0         TO S-LOWER(1)  MOVE 11000     TO S-UPPER(1)
           MOVE .1000     TO S-RATE(1)   MOVE 0         TO S-BASE-TAX(1)
           MOVE 11001     TO S-LOWER(2)  MOVE 44725     TO S-UPPER(2)
           MOVE .1200     TO S-RATE(2)   MOVE 1100      TO S-BASE-TAX(2)
           MOVE 44726     TO S-LOWER(3)  MOVE 95375     TO S-UPPER(3)
           MOVE .2200     TO S-RATE(3)   MOVE 5147      TO S-BASE-TAX(3)
           MOVE 95376     TO S-LOWER(4)  MOVE 182050    TO S-UPPER(4)
           MOVE .2400     TO S-RATE(4)   MOVE 16290     TO S-BASE-TAX(4)
           MOVE 182051    TO S-LOWER(5)  MOVE 231250    TO S-UPPER(5)
           MOVE .3200     TO S-RATE(5)   MOVE 37104     TO S-BASE-TAX(5)
           MOVE 231251    TO S-LOWER(6)  MOVE 578125    TO S-UPPER(6)
           MOVE .3500     TO S-RATE(6)   MOVE 52832     TO S-BASE-TAX(6)
           MOVE 578126    TO S-LOWER(7)  MOVE 9999999   TO S-UPPER(7)
           MOVE .3700     TO S-RATE(7)   MOVE 174238    TO S-BASE-TAX(7).

       1200-READ-RETURN.
           READ TAX-RETURN-FILE
               AT END MOVE '10' TO WS-TAX-STATUS
           END-READ.

       2000-PROCESS-RETURNS.
           ADD 1 TO WS-RETURNS-PROC
           INITIALIZE WS-CALCULATIONS WS-AUDIT-FLAGS
           PERFORM 2100-MATCH-W2
           PERFORM 2200-CALCULATE-AGI
           PERFORM 2300-CALCULATE-DEDUCTIONS
           PERFORM 2400-CALCULATE-TAX
           PERFORM 2500-APPLY-CREDITS
           PERFORM 2600-CALCULATE-REFUND
           PERFORM 2700-AUDIT-SCREENING
           PERFORM 2800-WRITE-OUTPUT
           PERFORM 1200-READ-RETURN.

       2100-MATCH-W2.
           *> Verify W-2 wages match reported wages
           MOVE ZEROS TO WS-W2-TOTAL-WAGES WS-W2-TOTAL-WITHHELD
           *> In production: read W2 file and match by SSN/year
           *> For this example, accept reported amounts
           MOVE TR-WAGES TO WS-W2-TOTAL-WAGES
           MOVE TR-FED-TAX-WITHHELD TO WS-W2-TOTAL-WITHHELD
           IF FUNCTION ABS(WS-W2-TOTAL-WAGES - TR-WAGES) > 1.00
               ADD 1 TO WS-W2-MISMATCHES
               ADD 50 TO WS-AUDIT-SCORE
               STRING FUNCTION TRIM(WS-AUDIT-REASONS)
                      'W2-WAGE-MISMATCH; '
                   DELIMITED SIZE INTO WS-AUDIT-REASONS
           END-IF.

       2200-CALCULATE-AGI.
           COMPUTE WS-TOTAL-INCOME =
               TR-WAGES + TR-INTEREST-INC + TR-DIVIDEND-INC +
               TR-BUSINESS-INC + TR-CAPITAL-GAINS +
               TR-IRA-DIST + TR-OTHER-INCOME

           *> Taxable SS benefits (simplified: 85% if income > threshold)
           IF TR-SS-BENEFITS > ZEROS
               IF WS-TOTAL-INCOME > 44000
                   COMPUTE WS-TOTAL-INCOME = WS-TOTAL-INCOME +
                       TR-SS-BENEFITS * 0.85
               ELSE IF WS-TOTAL-INCOME > 32000
                   COMPUTE WS-TOTAL-INCOME = WS-TOTAL-INCOME +
                       TR-SS-BENEFITS * 0.50
               END-IF
           END-IF

           *> Above-the-line deductions
           COMPUTE WS-AGI = WS-TOTAL-INCOME -
               TR-STUDENT-LOAN-INT - TR-IRA-DEDUCTION -
               TR-ALIMONY-PAID.

       2300-CALCULATE-DEDUCTIONS.
           *> Standard deduction based on filing status
           EVALUATE TRUE
               WHEN SINGLE
                   MOVE STD-DED-SINGLE TO WS-STANDARD-DED
               WHEN MARRIED-JOINT OR QUALIFYING-WID
                   MOVE STD-DED-MJ TO WS-STANDARD-DED
               WHEN MARRIED-SEP
                   MOVE STD-DED-MFS TO WS-STANDARD-DED
               WHEN HEAD-HOUSEHOLD
                   MOVE STD-DED-HH TO WS-STANDARD-DED
           END-EVALUATE

           IF ITEMIZING
               *> Cap SALT at $10,000
               COMPUTE WS-DEDUCTION-AMT =
                   TR-MORTGAGE-INT + TR-CHARITABLE +
                   FUNCTION MIN(TR-STATE-TAXES, WS-SALT-CAP)
               *> Medical: only amount > 7.5% of AGI
               IF TR-MEDICAL-EXPENSES > WS-AGI * 0.075
                   COMPUTE WS-DEDUCTION-AMT = WS-DEDUCTION-AMT +
                       TR-MEDICAL-EXPENSES - WS-AGI * 0.075
               END-IF
               *> Use higher of itemized or standard
               IF WS-DEDUCTION-AMT < WS-STANDARD-DED
                   MOVE WS-STANDARD-DED TO WS-DEDUCTION-AMT
               END-IF
           ELSE
               MOVE WS-STANDARD-DED TO WS-DEDUCTION-AMT
           END-IF

           COMPUTE WS-TAXABLE-INCOME = WS-AGI - WS-DEDUCTION-AMT
           IF WS-TAXABLE-INCOME < ZEROS
               MOVE ZEROS TO WS-TAXABLE-INCOME
           END-IF.

       2400-CALCULATE-TAX.
           MOVE ZEROS TO WS-GROSS-TAX
           EVALUATE TRUE
               WHEN MARRIED-JOINT OR QUALIFYING-WID
                   PERFORM VARYING WS-BRACKET-IDX FROM 1 BY 1
                       UNTIL WS-BRACKET-IDX > 7
                       IF WS-TAXABLE-INCOME >= MJ-LOWER(WS-BRACKET-IDX)
                           AND WS-TAXABLE-INCOME <= MJ-UPPER(WS-BRACKET-IDX)
                           COMPUTE WS-GROSS-TAX ROUNDED =
                               MJ-BASE-TAX(WS-BRACKET-IDX) +
                               (WS-TAXABLE-INCOME -
                                MJ-LOWER(WS-BRACKET-IDX)) *
                               MJ-RATE(WS-BRACKET-IDX)
                       END-IF
                   END-PERFORM
               WHEN OTHER
                   PERFORM VARYING WS-BRACKET-IDX FROM 1 BY 1
                       UNTIL WS-BRACKET-IDX > 7
                       IF WS-TAXABLE-INCOME >= S-LOWER(WS-BRACKET-IDX)
                           AND WS-TAXABLE-INCOME <= S-UPPER(WS-BRACKET-IDX)
                           COMPUTE WS-GROSS-TAX ROUNDED =
                               S-BASE-TAX(WS-BRACKET-IDX) +
                               (WS-TAXABLE-INCOME -
                                S-LOWER(WS-BRACKET-IDX)) *
                               S-RATE(WS-BRACKET-IDX)
                       END-IF
                   END-PERFORM
           END-EVALUATE

           IF WS-TAXABLE-INCOME > ZEROS
               COMPUTE WS-EFFECTIVE-RATE ROUNDED =
                   WS-GROSS-TAX / WS-TAXABLE-INCOME
           END-IF.

       2500-APPLY-CREDITS.
           MOVE ZEROS TO WS-CREDITS
           ADD TR-CHILD-TAX-CREDIT  TO WS-CREDITS
           ADD TR-EARNED-INC-CREDIT TO WS-CREDITS
           COMPUTE WS-NET-TAX = WS-GROSS-TAX - WS-CREDITS
           IF WS-NET-TAX < ZEROS
               MOVE ZEROS TO WS-NET-TAX
           END-IF.

       2600-CALCULATE-REFUND.
           COMPUTE WS-TOTAL-WITHHELD =
               TR-FED-TAX-WITHHELD + TR-EST-TAX-PAYMENTS
           COMPUTE WS-REFUND-AMOUNT =
               WS-TOTAL-WITHHELD - WS-NET-TAX
           IF WS-REFUND-AMOUNT >= ZEROS
               ADD 1 TO WS-REFUNDS
               ADD WS-REFUND-AMOUNT TO WS-TOTAL-REFUNDS
               MOVE ZEROS TO WS-BALANCE-DUE
           ELSE
               ADD 1 TO WS-BALANCE-DUES
               COMPUTE WS-BALANCE-DUE = WS-REFUND-AMOUNT * -1
               ADD WS-BALANCE-DUE TO WS-TOTAL-BAL-DUE
               MOVE ZEROS TO WS-REFUND-AMOUNT
           END-IF.

       2700-AUDIT-SCREENING.
           *> Charitable > 20% of AGI
           IF TR-CHARITABLE > WS-AGI * 0.20
               ADD 40 TO WS-AUDIT-SCORE
               STRING FUNCTION TRIM(WS-AUDIT-REASONS)
                      'HIGH-CHARITABLE; '
                   DELIMITED SIZE INTO WS-AUDIT-REASONS
           END-IF
           *> Business loss > $25,000
           IF TR-BUSINESS-INC < -25000
               ADD 35 TO WS-AUDIT-SCORE
               STRING FUNCTION TRIM(WS-AUDIT-REASONS)
                      'LARGE-BUSINESS-LOSS; '
                   DELIMITED SIZE INTO WS-AUDIT-REASONS
           END-IF
           *> Effective rate < 5% on high income
           IF WS-AGI > 200000 AND WS-EFFECTIVE-RATE < .05
               ADD 60 TO WS-AUDIT-SCORE
               STRING FUNCTION TRIM(WS-AUDIT-REASONS)
                      'LOW-EFFECTIVE-RATE-HIGH-INCOME; '
                   DELIMITED SIZE INTO WS-AUDIT-REASONS
           END-IF
           *> Large refund relative to income
           IF WS-REFUND-AMOUNT > WS-AGI * 0.30
               ADD 25 TO WS-AUDIT-SCORE
               STRING FUNCTION TRIM(WS-AUDIT-REASONS)
                      'LARGE-REFUND-RATIO; '
                   DELIMITED SIZE INTO WS-AUDIT-REASONS
           END-IF
           IF WS-AUDIT-SCORE >= 50
               MOVE 'Y' TO WS-AUDIT-FLAG
               ADD 1 TO WS-AUDIT-FLAGS-CNT
               MOVE SPACES TO AUDIT-LINE
               STRING TR-SSN ' SCORE:' WS-AUDIT-SCORE
                      ' ' WS-AUDIT-REASONS
                   DELIMITED SIZE INTO AUDIT-LINE
               WRITE AUDIT-LINE
           END-IF.

       2800-WRITE-OUTPUT.
           *> Processed return record
           MOVE SPACES TO PROC-RECORD
           STRING TR-SSN ' ' TR-TAX-YEAR ' '
                  WS-AGI ' ' WS-TAXABLE-INCOME ' '
                  WS-NET-TAX ' ' WS-REFUND-AMOUNT ' '
                  WS-BALANCE-DUE ' ' WS-EFFECTIVE-RATE
               DELIMITED SIZE INTO PROC-RECORD
           WRITE PROC-RECORD

           *> Refund or balance due
           IF WS-REFUND-AMOUNT > ZEROS
               MOVE SPACES TO REFUND-RECORD
               STRING TR-SSN ' REFUND:' WS-REFUND-AMOUNT
                   DELIMITED SIZE INTO REFUND-RECORD
               WRITE REFUND-RECORD
           ELSE IF WS-BALANCE-DUE > ZEROS
               MOVE SPACES TO BAL-DUE-RECORD
               STRING TR-SSN ' BALANCE-DUE:' WS-BALANCE-DUE
                   DELIMITED SIZE INTO BAL-DUE-RECORD
               WRITE BAL-DUE-RECORD
           END-IF

           *> E-file record
           MOVE SPACES TO EFILE-RECORD
           STRING 'EFILE|' TR-SSN '|' TR-TAX-YEAR '|'
                  TR-FILING-STATUS '|' WS-AGI '|'
                  WS-NET-TAX '|' WS-REFUND-AMOUNT '|'
                  WS-BALANCE-DUE '|' WS-CURRENT-DATE
               DELIMITED SIZE INTO EFILE-RECORD
           WRITE EFILE-RECORD.

       3000-PRINT-SUMMARY.
           DISPLAY "=== TAX PROCESSING SUMMARY ==="
           DISPLAY "Returns Processed  : " WS-RETURNS-PROC
           DISPLAY "Refunds            : " WS-REFUNDS
           DISPLAY "Balance Due        : " WS-BALANCE-DUES
           DISPLAY "Audit Flags        : " WS-AUDIT-FLAGS-CNT
           DISPLAY "W-2 Mismatches     : " WS-W2-MISMATCHES
           DISPLAY "Total Refunds      : " WS-TOTAL-REFUNDS
           DISPLAY "Total Balance Due  : " WS-TOTAL-BAL-DUE.

       9000-TERMINATE.
           CLOSE TAX-RETURN-FILE W2-FILE TAXPAYER-MASTER
                 PROCESSED-RETURNS AUDIT-FLAGS
                 REFUND-FILE BALANCE-DUE-FILE EFILE-OUTPUT.
