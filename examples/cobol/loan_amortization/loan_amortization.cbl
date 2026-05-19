      *> ============================================================
      *> LOAN AMORTIZATION AND MORTGAGE PROCESSING SYSTEM
      *> ============================================================
      *> Full mortgage lifecycle: origination, amortization schedule
      *> generation, payment processing, escrow management,
      *> payoff calculations, late fee assessment, 1098 tax forms.
      *>
      *> Demonstrates: COMPUTE with complex financial formulas,
      *> PERFORM VARYING with multiple conditions, EVALUATE TRUE,
      *> packed decimal arithmetic, date arithmetic using
      *> intrinsic functions, multi-level report generation.
      *> ============================================================
       IDENTIFICATION DIVISION.
       PROGRAM-ID. LOAN-AMORTIZATION.

       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT LOAN-MASTER ASSIGN TO "loans.idx"
               ORGANIZATION IS INDEXED
               ACCESS MODE IS DYNAMIC
               RECORD KEY IS LN-LOAN-NUMBER
               FILE STATUS IS WS-LN-STATUS.

           SELECT PAYMENT-FILE ASSIGN TO "payments.dat"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-PAY-STATUS.

           SELECT AMORT-SCHEDULE ASSIGN TO "amortization.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT PAYMENT-HISTORY ASSIGN TO "payment_history.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT TAX-FORM-1098 ASSIGN TO "form_1098.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT DELINQUENCY-RPT ASSIGN TO "delinquency.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

       DATA DIVISION.
       FILE SECTION.

       FD  LOAN-MASTER
           RECORD CONTAINS 400 CHARACTERS.
       01  LOAN-RECORD.
           05  LN-LOAN-NUMBER      PIC X(15).
           05  LN-BORROWER-NAME    PIC X(50).
           05  LN-BORROWER-SSN     PIC X(11).
           05  LN-CO-BORROWER      PIC X(50).
           05  LN-PROPERTY-ADDR    PIC X(80).
           05  LN-LOAN-TYPE        PIC X(4).
               88  CONVENTIONAL    VALUE 'CONV'.
               88  FHA-LOAN        VALUE 'FHA '.
               88  VA-LOAN         VALUE 'VA  '.
               88  JUMBO-LOAN      VALUE 'JUMB'.
           05  LN-ORIGINAL-AMOUNT  PIC 9(11)V99 COMP-3.
           05  LN-CURRENT-BALANCE  PIC 9(11)V99 COMP-3.
           05  LN-INTEREST-RATE    PIC 9(2)V9(6) COMP-3.
           05  LN-TERM-MONTHS      PIC 9(4).
           05  LN-REMAINING-MONTHS PIC 9(4).
           05  LN-ORIGINATION-DATE PIC 9(8).
           05  LN-FIRST-PMT-DATE   PIC 9(8).
           05  LN-MATURITY-DATE    PIC 9(8).
           05  LN-MONTHLY-PAYMENT  PIC 9(9)V99 COMP-3.
           05  LN-ESCROW-PAYMENT   PIC 9(7)V99 COMP-3.
           05  LN-TOTAL-PAYMENT    PIC 9(9)V99 COMP-3.
           05  LN-ESCROW-BALANCE   PIC 9(9)V99 COMP-3.
           05  LN-ANNUAL-TAX       PIC 9(9)V99 COMP-3.
           05  LN-ANNUAL-INSURANCE PIC 9(7)V99 COMP-3.
           05  LN-PMI-AMOUNT       PIC 9(7)V99 COMP-3.
           05  LN-LOAN-STATUS      PIC X(2).
               88  LN-CURRENT      VALUE 'CU'.
               88  LN-30-DAY       VALUE '30'.
               88  LN-60-DAY       VALUE '60'.
               88  LN-90-DAY       VALUE '90'.
               88  LN-FORECLOSURE  VALUE 'FC'.
               88  LN-PAID-OFF     VALUE 'PO'.
           05  LN-LAST-PMT-DATE    PIC 9(8).
           05  LN-LAST-PMT-AMOUNT  PIC 9(9)V99 COMP-3.
           05  LN-DAYS-DELINQUENT  PIC 9(4).
           05  LN-YTD-INTEREST     PIC 9(11)V99 COMP-3.
           05  LN-YTD-PRINCIPAL    PIC 9(11)V99 COMP-3.
           05  LN-YTD-ESCROW       PIC 9(9)V99 COMP-3.
           05  LN-LIFE-INTEREST    PIC 9(13)V99 COMP-3.
           05  LN-LIFE-PRINCIPAL   PIC 9(11)V99 COMP-3.
           05  LN-LATE-CHARGES     PIC 9(9)V99 COMP-3.
           05  LN-LATE-CHARGE-RATE PIC V9(4) COMP-3.
           05  FILLER              PIC X(20).

       FD  PAYMENT-FILE
           RECORD CONTAINS 100 CHARACTERS.
       01  PAYMENT-RECORD.
           05  PMT-LOAN-NUMBER     PIC X(15).
           05  PMT-DATE            PIC 9(8).
           05  PMT-AMOUNT          PIC 9(9)V99.
           05  PMT-TYPE            PIC X(2).
               88  PMT-REGULAR     VALUE 'RE'.
               88  PMT-EXTRA-PRIN  VALUE 'EP'.
               88  PMT-ESCROW-ONLY VALUE 'ES'.
               88  PMT-PAYOFF      VALUE 'PO'.
           05  PMT-CHECK-NUMBER    PIC X(10).
           05  FILLER              PIC X(55).

       FD  AMORT-SCHEDULE
           RECORD CONTAINS 132 CHARACTERS.
       01  AMORT-LINE              PIC X(132).

       FD  PAYMENT-HISTORY
           RECORD CONTAINS 132 CHARACTERS.
       01  HIST-LINE               PIC X(132).

       FD  TAX-FORM-1098
           RECORD CONTAINS 132 CHARACTERS.
       01  FORM-LINE               PIC X(132).

       FD  DELINQUENCY-RPT
           RECORD CONTAINS 132 CHARACTERS.
       01  DELINQ-LINE             PIC X(132).

       WORKING-STORAGE SECTION.

       01  WS-STATUS.
           05  WS-LN-STATUS        PIC XX.
               88  LN-OK           VALUE '00'.
               88  LN-NOT-FOUND    VALUE '23'.
               88  LN-EOF          VALUE '10'.
           05  WS-PAY-STATUS       PIC XX.
               88  PAY-OK          VALUE '00'.
               88  PAY-EOF         VALUE '10'.

       01  WS-FINANCIAL-CALC.
           05  WS-MONTHLY-RATE     PIC 9(2)V9(10) COMP-3.
           05  WS-INTEREST-PORTION PIC 9(9)V99    COMP-3.
           05  WS-PRINCIPAL-PORTION PIC 9(9)V99   COMP-3.
           05  WS-ESCROW-PORTION   PIC 9(7)V99    COMP-3.
           05  WS-NEW-BALANCE      PIC 9(11)V99   COMP-3.
           05  WS-PAYOFF-AMOUNT    PIC 9(11)V99   COMP-3.
           05  WS-LATE-FEE         PIC 9(7)V99    COMP-3.
           05  WS-OVERPAYMENT      PIC 9(9)V99    COMP-3.
           05  WS-PAYMENT-DUE-DATE PIC 9(8).
           05  WS-DAYS-LATE        PIC 9(4).
           05  WS-GRACE-PERIOD     PIC 9(2) VALUE 15.

       01  WS-AMORT-WORK.
           05  WS-AMORT-BALANCE    PIC 9(11)V99 COMP-3.
           05  WS-AMORT-MONTH      PIC 9(4).
           05  WS-TOTAL-INTEREST   PIC 9(13)V99 COMP-3.
           05  WS-TOTAL-PRINCIPAL  PIC 9(11)V99 COMP-3.
           05  WS-AMORT-DATE       PIC 9(8).

       01  WS-COUNTERS.
           05  WS-LOANS-PROCESSED  PIC 9(8) VALUE ZEROS.
           05  WS-PAYMENTS-POSTED  PIC 9(8) VALUE ZEROS.
           05  WS-LATE-FEES-ASSESSED PIC 9(6) VALUE ZEROS.
           05  WS-DELINQUENT-COUNT PIC 9(6) VALUE ZEROS.
           05  WS-TOTAL-PORTFOLIO  PIC 9(15)V99 VALUE ZEROS.
           05  WS-TOTAL-LATE-FEES  PIC 9(11)V99 VALUE ZEROS.

       01  WS-FORMATTED.
           05  WF-AMOUNT           PIC ZZZ,ZZZ,ZZZ,ZZ9.99.
           05  WF-RATE             PIC Z9.9999.
           05  WF-MONTHS           PIC ZZZ9.
           05  WF-DATE             PIC 9999/99/99.

       01  WS-CURRENT-DATE         PIC 9(8).

       PROCEDURE DIVISION.

       0000-MAIN.
           PERFORM 1000-INITIALIZE
           PERFORM 2000-PROCESS-PAYMENTS
               UNTIL PAY-EOF
           PERFORM 3000-ASSESS-LATE-FEES
           PERFORM 4000-GENERATE-AMORTIZATION
           PERFORM 5000-GENERATE-1098-FORMS
           PERFORM 6000-GENERATE-DELINQUENCY
           PERFORM 7000-PRINT-SUMMARY
           PERFORM 9000-TERMINATE
           STOP RUN.

       1000-INITIALIZE.
           MOVE FUNCTION CURRENT-DATE(1:8) TO WS-CURRENT-DATE
           OPEN I-O    LOAN-MASTER
           OPEN INPUT  PAYMENT-FILE
           OPEN OUTPUT AMORT-SCHEDULE
           OPEN OUTPUT PAYMENT-HISTORY
           OPEN OUTPUT TAX-FORM-1098
           OPEN OUTPUT DELINQUENCY-RPT
           PERFORM 1100-READ-PAYMENT.

       1100-READ-PAYMENT.
           READ PAYMENT-FILE
               AT END MOVE '10' TO WS-PAY-STATUS
           END-READ.

       2000-PROCESS-PAYMENTS.
           MOVE PMT-LOAN-NUMBER TO LN-LOAN-NUMBER
           READ LOAN-MASTER
               INVALID KEY
                   ADD 1 TO WS-PAYMENTS-POSTED
                   PERFORM 1100-READ-PAYMENT
                   STOP RUN
           END-READ
           IF LN-OK
               PERFORM 2100-CALCULATE-PAYMENT-SPLIT
               PERFORM 2200-APPLY-PAYMENT
               PERFORM 2300-UPDATE-LOAN
               PERFORM 2400-WRITE-HISTORY
               ADD 1 TO WS-PAYMENTS-POSTED
           END-IF
           PERFORM 1100-READ-PAYMENT.

       2100-CALCULATE-PAYMENT-SPLIT.
           *> Monthly interest rate
           COMPUTE WS-MONTHLY-RATE =
               LN-INTEREST-RATE / 12

           *> Interest portion = balance * monthly rate
           COMPUTE WS-INTEREST-PORTION ROUNDED =
               LN-CURRENT-BALANCE * WS-MONTHLY-RATE

           *> Check for late payment
           COMPUTE WS-DAYS-LATE =
               FUNCTION INTEGER-OF-DATE(PMT-DATE) -
               FUNCTION INTEGER-OF-DATE(LN-LAST-PMT-DATE) - 30
           IF WS-DAYS-LATE > WS-GRACE-PERIOD
               COMPUTE WS-LATE-FEE ROUNDED =
                   LN-MONTHLY-PAYMENT * LN-LATE-CHARGE-RATE
               ADD WS-LATE-FEE TO LN-LATE-CHARGES
               ADD 1 TO WS-LATE-FEES-ASSESSED
               ADD WS-LATE-FEE TO WS-TOTAL-LATE-FEES
           ELSE
               MOVE ZEROS TO WS-LATE-FEE
           END-IF

           *> Escrow portion
           MOVE LN-ESCROW-PAYMENT TO WS-ESCROW-PORTION

           *> Principal = payment - interest - escrow - late fee
           COMPUTE WS-PRINCIPAL-PORTION =
               PMT-AMOUNT - WS-INTEREST-PORTION -
               WS-ESCROW-PORTION - WS-LATE-FEE

           IF WS-PRINCIPAL-PORTION < ZEROS
               MOVE ZEROS TO WS-PRINCIPAL-PORTION
           END-IF.

       2200-APPLY-PAYMENT.
           EVALUATE TRUE
               WHEN PMT-PAYOFF
                   *> Full payoff: calculate exact payoff amount
                   COMPUTE WS-PAYOFF-AMOUNT =
                       LN-CURRENT-BALANCE + WS-INTEREST-PORTION +
                       LN-LATE-CHARGES
                   MOVE ZEROS TO LN-CURRENT-BALANCE
                   MOVE 'PO' TO LN-LOAN-STATUS
               WHEN PMT-EXTRA-PRIN
                   *> Extra principal payment
                   SUBTRACT PMT-AMOUNT FROM LN-CURRENT-BALANCE
                   ADD PMT-AMOUNT TO WS-PRINCIPAL-PORTION
               WHEN OTHER
                   *> Regular payment
                   SUBTRACT WS-PRINCIPAL-PORTION FROM LN-CURRENT-BALANCE
                   ADD WS-ESCROW-PORTION TO LN-ESCROW-BALANCE
           END-EVALUATE

           *> Update YTD accumulators
           ADD WS-INTEREST-PORTION  TO LN-YTD-INTEREST
           ADD WS-PRINCIPAL-PORTION TO LN-YTD-PRINCIPAL
           ADD WS-ESCROW-PORTION    TO LN-YTD-ESCROW

           *> Update life-of-loan accumulators
           ADD WS-INTEREST-PORTION  TO LN-LIFE-INTEREST
           ADD WS-PRINCIPAL-PORTION TO LN-LIFE-PRINCIPAL

           *> Update LTV — remove PMI if balance < 80% of original
           IF LN-CURRENT-BALANCE < LN-ORIGINAL-AMOUNT * 0.80
               MOVE ZEROS TO LN-PMI-AMOUNT
           END-IF

           SUBTRACT 1 FROM LN-REMAINING-MONTHS.

       2300-UPDATE-LOAN.
           MOVE PMT-DATE   TO LN-LAST-PMT-DATE
           MOVE PMT-AMOUNT TO LN-LAST-PMT-AMOUNT

           *> Update delinquency status
           EVALUATE TRUE
               WHEN WS-DAYS-LATE <= 0
                   MOVE 'CU' TO LN-LOAN-STATUS
                   MOVE ZEROS TO LN-DAYS-DELINQUENT
               WHEN WS-DAYS-LATE <= 30
                   MOVE '30' TO LN-LOAN-STATUS
                   MOVE WS-DAYS-LATE TO LN-DAYS-DELINQUENT
               WHEN WS-DAYS-LATE <= 60
                   MOVE '60' TO LN-LOAN-STATUS
                   MOVE WS-DAYS-LATE TO LN-DAYS-DELINQUENT
               WHEN WS-DAYS-LATE <= 90
                   MOVE '90' TO LN-LOAN-STATUS
                   MOVE WS-DAYS-LATE TO LN-DAYS-DELINQUENT
               WHEN OTHER
                   MOVE 'FC' TO LN-LOAN-STATUS
                   MOVE WS-DAYS-LATE TO LN-DAYS-DELINQUENT
           END-EVALUATE

           REWRITE LOAN-RECORD
               INVALID KEY CONTINUE
           END-REWRITE.

       2400-WRITE-HISTORY.
           MOVE LN-CURRENT-BALANCE TO WF-AMOUNT
           MOVE SPACES TO HIST-LINE
           STRING LN-LOAN-NUMBER ' '
                  PMT-DATE ' '
                  PMT-AMOUNT ' INT:'
                  WS-INTEREST-PORTION ' PRIN:'
                  WS-PRINCIPAL-PORTION ' BAL:'
                  WF-AMOUNT
               DELIMITED SIZE INTO HIST-LINE
           WRITE HIST-LINE.

       3000-ASSESS-LATE-FEES.
           *> Scan all active loans for delinquency
           MOVE LOW-VALUES TO LN-LOAN-NUMBER
           START LOAN-MASTER KEY >= LN-LOAN-NUMBER
               INVALID KEY STOP RUN
           END-START
           PERFORM 3100-LATE-FEE-SCAN UNTIL LN-EOF.

       3100-LATE-FEE-SCAN.
           READ LOAN-MASTER NEXT
               AT END MOVE '10' TO WS-LN-STATUS
           END-READ
           IF NOT LN-EOF
               IF LN-DAYS-DELINQUENT > WS-GRACE-PERIOD
                   ADD 1 TO WS-DELINQUENT-COUNT
               END-IF
               ADD LN-CURRENT-BALANCE TO WS-TOTAL-PORTFOLIO
           END-IF.

       4000-GENERATE-AMORTIZATION.
           *> Generate schedule for first loan found
           MOVE LOW-VALUES TO LN-LOAN-NUMBER
           START LOAN-MASTER KEY >= LN-LOAN-NUMBER
               INVALID KEY STOP RUN
           END-START
           READ LOAN-MASTER NEXT
               AT END STOP RUN
           END-READ

           WRITE AMORT-LINE FROM ALL '='
           MOVE SPACES TO AMORT-LINE
           STRING 'AMORTIZATION SCHEDULE: LOAN ' LN-LOAN-NUMBER
                  '  BORROWER: ' LN-BORROWER-NAME
               DELIMITED SIZE INTO AMORT-LINE
           WRITE AMORT-LINE
           MOVE LN-INTEREST-RATE TO WF-RATE
           MOVE LN-TERM-MONTHS   TO WF-MONTHS
           MOVE SPACES TO AMORT-LINE
           STRING 'Original Amount: ' LN-ORIGINAL-AMOUNT
                  '  Rate: ' WF-RATE '%'
                  '  Term: ' WF-MONTHS ' months'
               DELIMITED SIZE INTO AMORT-LINE
           WRITE AMORT-LINE
           WRITE AMORT-LINE FROM ALL '-'
           WRITE AMORT-LINE FROM
               "PMT#  PAYMENT     INTEREST    PRINCIPAL   " &
               "ESCROW      BALANCE     CUMUL-INT"
           WRITE AMORT-LINE FROM ALL '-'

           MOVE LN-ORIGINAL-AMOUNT TO WS-AMORT-BALANCE
           COMPUTE WS-MONTHLY-RATE = LN-INTEREST-RATE / 12
           MOVE ZEROS TO WS-TOTAL-INTEREST WS-TOTAL-PRINCIPAL

           PERFORM VARYING WS-AMORT-MONTH FROM 1 BY 1
               UNTIL WS-AMORT-MONTH > LN-TERM-MONTHS
               COMPUTE WS-INTEREST-PORTION ROUNDED =
                   WS-AMORT-BALANCE * WS-MONTHLY-RATE
               COMPUTE WS-PRINCIPAL-PORTION =
                   LN-MONTHLY-PAYMENT - WS-INTEREST-PORTION
               SUBTRACT WS-PRINCIPAL-PORTION FROM WS-AMORT-BALANCE
               ADD WS-INTEREST-PORTION  TO WS-TOTAL-INTEREST
               ADD WS-PRINCIPAL-PORTION TO WS-TOTAL-PRINCIPAL

               IF WS-AMORT-BALANCE < ZEROS
                   MOVE ZEROS TO WS-AMORT-BALANCE
               END-IF

               MOVE SPACES TO AMORT-LINE
               STRING WS-AMORT-MONTH ' '
                      LN-MONTHLY-PAYMENT ' '
                      WS-INTEREST-PORTION ' '
                      WS-PRINCIPAL-PORTION ' '
                      LN-ESCROW-PAYMENT ' '
                      WS-AMORT-BALANCE ' '
                      WS-TOTAL-INTEREST
                   DELIMITED SIZE INTO AMORT-LINE
               WRITE AMORT-LINE
           END-PERFORM

           WRITE AMORT-LINE FROM ALL '='
           MOVE SPACES TO AMORT-LINE
           STRING 'TOTAL INTEREST PAID: ' WS-TOTAL-INTEREST
               DELIMITED SIZE INTO AMORT-LINE
           WRITE AMORT-LINE.

       5000-GENERATE-1098-FORMS.
           MOVE LOW-VALUES TO LN-LOAN-NUMBER
           START LOAN-MASTER KEY >= LN-LOAN-NUMBER
               INVALID KEY STOP RUN
           END-START
           WRITE FORM-LINE FROM
               "MORTGAGE INTEREST STATEMENTS - FORM 1098"
           WRITE FORM-LINE FROM ALL '='
           PERFORM 5100-1098-SCAN UNTIL LN-EOF.

       5100-1098-SCAN.
           READ LOAN-MASTER NEXT
               AT END MOVE '10' TO WS-LN-STATUS
           END-READ
           IF NOT LN-EOF AND LN-YTD-INTEREST > ZEROS
               MOVE SPACES TO FORM-LINE
               STRING 'BORROWER: ' LN-BORROWER-NAME
                      '  SSN: ' LN-BORROWER-SSN
                   DELIMITED SIZE INTO FORM-LINE
               WRITE FORM-LINE
               MOVE SPACES TO FORM-LINE
               STRING '  Mortgage Interest: ' LN-YTD-INTEREST
                      '  Points Paid: 0.00'
                      '  Refund: 0.00'
                   DELIMITED SIZE INTO FORM-LINE
               WRITE FORM-LINE
               MOVE SPACES TO FORM-LINE
               STRING '  Property: ' LN-PROPERTY-ADDR
                   DELIMITED SIZE INTO FORM-LINE
               WRITE FORM-LINE
               WRITE FORM-LINE FROM ALL '-'
           END-IF.

       6000-GENERATE-DELINQUENCY.
           MOVE LOW-VALUES TO LN-LOAN-NUMBER
           START LOAN-MASTER KEY >= LN-LOAN-NUMBER
               INVALID KEY STOP RUN
           END-START
           WRITE DELINQ-LINE FROM "DELINQUENCY REPORT"
           WRITE DELINQ-LINE FROM ALL '='
           WRITE DELINQ-LINE FROM
               "LOAN NUMBER     BORROWER                    " &
               "BALANCE         DAYS LATE  STATUS"
           PERFORM 6100-DELINQ-SCAN UNTIL LN-EOF.

       6100-DELINQ-SCAN.
           READ LOAN-MASTER NEXT
               AT END MOVE '10' TO WS-LN-STATUS
           END-READ
           IF NOT LN-EOF AND LN-DAYS-DELINQUENT > ZEROS
               MOVE LN-CURRENT-BALANCE TO WF-AMOUNT
               MOVE SPACES TO DELINQ-LINE
               STRING LN-LOAN-NUMBER ' '
                      LN-BORROWER-NAME ' '
                      WF-AMOUNT ' '
                      LN-DAYS-DELINQUENT ' '
                      LN-LOAN-STATUS
                   DELIMITED SIZE INTO DELINQ-LINE
               WRITE DELINQ-LINE
           END-IF.

       7000-PRINT-SUMMARY.
           DISPLAY "=== LOAN PROCESSING SUMMARY ==="
           DISPLAY "Payments Posted    : " WS-PAYMENTS-POSTED
           DISPLAY "Late Fees Assessed : " WS-LATE-FEES-ASSESSED
           DISPLAY "Total Late Fees    : " WS-TOTAL-LATE-FEES
           DISPLAY "Delinquent Loans   : " WS-DELINQUENT-COUNT
           DISPLAY "Total Portfolio    : " WS-TOTAL-PORTFOLIO.

       9000-TERMINATE.
           CLOSE LOAN-MASTER PAYMENT-FILE AMORT-SCHEDULE
                 PAYMENT-HISTORY TAX-FORM-1098 DELINQUENCY-RPT.
