      *> ============================================================
      *> INSURANCE CLAIMS PROCESSING SYSTEM
      *> ============================================================
      *> Adjudicates medical insurance claims: eligibility check,
      *> benefit calculation, deductible tracking, EOB generation.
      *>
      *> Demonstrates: XML GENERATE/PARSE (COBOL 2014+),
      *> JSON GENERATE (COBOL 2014+), VALIDATE statement,
      *> FUNCTION intrinsics, nested EVALUATE, 
      *> complex group moves, CORRESPONDING.
      *> ============================================================
       IDENTIFICATION DIVISION.
       PROGRAM-ID. INSURANCE-CLAIMS.

       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT CLAIMS-FILE ASSIGN TO "claims_input.dat"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-CLM-STATUS.

           SELECT MEMBER-FILE ASSIGN TO "members.dat"
               ORGANIZATION IS INDEXED
               ACCESS MODE IS RANDOM
               RECORD KEY IS MBR-ID
               FILE STATUS IS WS-MBR-STATUS.

           SELECT PROVIDER-FILE ASSIGN TO "providers.dat"
               ORGANIZATION IS INDEXED
               ACCESS MODE IS RANDOM
               RECORD KEY IS PRV-NPI
               FILE STATUS IS WS-PRV-STATUS.

           SELECT EOB-OUTPUT ASSIGN TO "eob_output.txt"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-EOB-STATUS.

           SELECT CLAIMS-PAID ASSIGN TO "claims_paid.dat"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-PAID-STATUS.

           SELECT CLAIMS-DENIED ASSIGN TO "claims_denied.dat"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-DENY-STATUS.

       DATA DIVISION.
       FILE SECTION.

       FD  CLAIMS-FILE
           RECORD CONTAINS 300 CHARACTERS.
       01  CLAIM-RECORD.
           05  CLM-CLAIM-ID        PIC X(15).
           05  CLM-MEMBER-ID       PIC X(12).
           05  CLM-PROVIDER-NPI    PIC X(10).
           05  CLM-SERVICE-DATE    PIC 9(8).
           05  CLM-RECEIVED-DATE   PIC 9(8).
           05  CLM-CLAIM-TYPE      PIC X(2).
               88  CLM-MEDICAL     VALUE 'ME'.
               88  CLM-DENTAL      VALUE 'DE'.
               88  CLM-VISION      VALUE 'VI'.
               88  CLM-PHARMACY    VALUE 'PH'.
               88  CLM-MENTAL-HLTH VALUE 'MH'.
           05  CLM-DIAGNOSIS-CODE  PIC X(8).
           05  CLM-PROCEDURE-CODE  PIC X(8).
           05  CLM-BILLED-AMOUNT   PIC 9(9)V99.
           05  CLM-UNITS           PIC 9(5)V99.
           05  CLM-PLACE-OF-SVC    PIC X(2).
               88  CLM-INPATIENT   VALUE '21'.
               88  CLM-OUTPATIENT  VALUE '22'.
               88  CLM-OFFICE      VALUE '11'.
               88  CLM-EMERGENCY   VALUE '23'.
               88  CLM-TELEHEALTH  VALUE '02'.
           05  CLM-MODIFIER        PIC X(8).
           05  CLM-PRIOR-AUTH      PIC X(12).
           05  FILLER              PIC X(148).

       FD  MEMBER-FILE
           RECORD CONTAINS 250 CHARACTERS.
       01  MEMBER-RECORD.
           05  MBR-ID              PIC X(12).
           05  MBR-NAME            PIC X(40).
           05  MBR-DOB             PIC 9(8).
           05  MBR-PLAN-CODE       PIC X(6).
           05  MBR-EFFECTIVE-DATE  PIC 9(8).
           05  MBR-TERM-DATE       PIC 9(8).
           05  MBR-STATUS          PIC X(1).
               88  MBR-ACTIVE      VALUE 'A'.
               88  MBR-TERMED      VALUE 'T'.
               88  MBR-COBRA       VALUE 'C'.
           05  MBR-DEDUCTIBLE-IND  PIC 9(7)V99.
           05  MBR-DEDUCTIBLE-MET  PIC 9(7)V99.
           05  MBR-OOP-MAX         PIC 9(7)V99.
           05  MBR-OOP-MET         PIC 9(7)V99.
           05  MBR-COPAY-OFFICE    PIC 9(5)V99.
           05  MBR-COPAY-SPEC      PIC 9(5)V99.
           05  MBR-COPAY-ER        PIC 9(5)V99.
           05  MBR-COINSURANCE-PCT PIC 9(3)V99.
           05  MBR-GROUP-NUM       PIC X(10).
           05  FILLER              PIC X(80).

       FD  PROVIDER-FILE
           RECORD CONTAINS 150 CHARACTERS.
       01  PROVIDER-RECORD.
           05  PRV-NPI             PIC X(10).
           05  PRV-NAME            PIC X(50).
           05  PRV-SPECIALTY       PIC X(6).
           05  PRV-NETWORK-STATUS  PIC X(1).
               88  PRV-IN-NETWORK  VALUE 'I'.
               88  PRV-OUT-NETWORK VALUE 'O'.
           05  PRV-CONTRACT-RATE   PIC V9(4).
           05  FILLER              PIC X(83).

       FD  EOB-OUTPUT
           RECORD CONTAINS 132 CHARACTERS.
       01  EOB-LINE                PIC X(132).

       FD  CLAIMS-PAID
           RECORD CONTAINS 200 CHARACTERS.
       01  PAID-RECORD             PIC X(200).

       FD  CLAIMS-DENIED
           RECORD CONTAINS 200 CHARACTERS.
       01  DENIED-RECORD           PIC X(200).

       WORKING-STORAGE SECTION.

       01  WS-STATUS.
           05  WS-CLM-STATUS       PIC XX.
               88  CLM-OK          VALUE '00'.
               88  CLM-EOF         VALUE '10'.
           05  WS-MBR-STATUS       PIC XX.
               88  MBR-FOUND       VALUE '00'.
               88  MBR-NOT-FOUND   VALUE '23'.
           05  WS-PRV-STATUS       PIC XX.
               88  PRV-FOUND       VALUE '00'.
               88  PRV-NOT-FOUND   VALUE '23'.
           05  WS-EOB-STATUS       PIC XX.
           05  WS-PAID-STATUS      PIC XX.
           05  WS-DENY-STATUS      PIC XX.

       01  WS-ADJUDICATION.
           05  WS-ALLOWED-AMOUNT   PIC 9(9)V99 VALUE ZEROS.
           05  WS-DEDUCTIBLE-APPLY PIC 9(7)V99 VALUE ZEROS.
           05  WS-COINSURANCE-AMT  PIC 9(7)V99 VALUE ZEROS.
           05  WS-COPAY-AMOUNT     PIC 9(5)V99 VALUE ZEROS.
           05  WS-PLAN-PAYS        PIC 9(9)V99 VALUE ZEROS.
           05  WS-MEMBER-PAYS      PIC 9(9)V99 VALUE ZEROS.
           05  WS-NOT-COVERED      PIC 9(9)V99 VALUE ZEROS.
           05  WS-DENIAL-REASON    PIC X(60) VALUE SPACES.
           05  WS-CLAIM-STATUS     PIC X(1).
               88  CLAIM-APPROVED  VALUE 'A'.
               88  CLAIM-DENIED    VALUE 'D'.
               88  CLAIM-PENDED    VALUE 'P'.

       01  WS-BENEFIT-TABLE.
           05  BEN-ENTRY OCCURS 10 TIMES INDEXED BY BEN-IDX.
               10  BEN-PROC-PREFIX PIC X(3).
               10  BEN-COVERED     PIC X(1).
               10  BEN-REQUIRES-AUTH PIC X(1).
               10  BEN-MAX-UNITS   PIC 9(5)V99.
               10  BEN-ALLOWED-PCT PIC V9(4).

       01  WS-COUNTERS.
           05  WS-CLAIMS-PROCESSED PIC 9(8) VALUE ZEROS.
           05  WS-CLAIMS-APPROVED  PIC 9(8) VALUE ZEROS.
           05  WS-CLAIMS-DENIED    PIC 9(8) VALUE ZEROS.
           05  WS-CLAIMS-PENDED    PIC 9(8) VALUE ZEROS.
           05  WS-TOTAL-BILLED     PIC 9(13)V99 VALUE ZEROS.
           05  WS-TOTAL-ALLOWED    PIC 9(13)V99 VALUE ZEROS.
           05  WS-TOTAL-PLAN-PAYS  PIC 9(13)V99 VALUE ZEROS.
           05  WS-TOTAL-MBR-PAYS   PIC 9(13)V99 VALUE ZEROS.

       01  WS-CURRENT-DATE         PIC 9(8).

       01  WS-EOB-FIELDS.
           05  WF-BILLED           PIC ZZZ,ZZZ,ZZ9.99.
           05  WF-ALLOWED          PIC ZZZ,ZZZ,ZZ9.99.
           05  WF-DEDUCTIBLE       PIC ZZZ,ZZZ,ZZ9.99.
           05  WF-COINSURANCE      PIC ZZZ,ZZZ,ZZ9.99.
           05  WF-COPAY            PIC ZZZ,ZZZ,ZZ9.99.
           05  WF-PLAN-PAYS        PIC ZZZ,ZZZ,ZZ9.99.
           05  WF-MBR-PAYS         PIC ZZZ,ZZZ,ZZ9.99.

       PROCEDURE DIVISION.

       0000-MAIN.
           PERFORM 1000-INITIALIZE
           PERFORM 2000-PROCESS-CLAIMS
               UNTIL CLM-EOF
           PERFORM 3000-PRINT-SUMMARY
           PERFORM 9000-TERMINATE
           STOP RUN.

       1000-INITIALIZE.
           MOVE FUNCTION CURRENT-DATE(1:8) TO WS-CURRENT-DATE
           OPEN INPUT  CLAIMS-FILE
           OPEN I-O    MEMBER-FILE
           OPEN INPUT  PROVIDER-FILE
           OPEN OUTPUT EOB-OUTPUT
           OPEN OUTPUT CLAIMS-PAID
           OPEN OUTPUT CLAIMS-DENIED
           PERFORM 1100-LOAD-BENEFIT-TABLE
           PERFORM 1200-READ-CLAIM.

       1100-LOAD-BENEFIT-TABLE.
           MOVE '990' TO BEN-PROC-PREFIX(1)
           MOVE 'Y'   TO BEN-COVERED(1)
           MOVE 'N'   TO BEN-REQUIRES-AUTH(1)
           MOVE 1.00  TO BEN-ALLOWED-PCT(1)
           MOVE '992' TO BEN-PROC-PREFIX(2)
           MOVE 'Y'   TO BEN-COVERED(2)
           MOVE 'N'   TO BEN-REQUIRES-AUTH(2)
           MOVE 1.00  TO BEN-ALLOWED-PCT(2)
           MOVE '993' TO BEN-PROC-PREFIX(3)
           MOVE 'Y'   TO BEN-COVERED(3)
           MOVE 'Y'   TO BEN-REQUIRES-AUTH(3)
           MOVE 0.80  TO BEN-ALLOWED-PCT(3)
           MOVE '270' TO BEN-PROC-PREFIX(4)
           MOVE 'Y'   TO BEN-COVERED(4)
           MOVE 'Y'   TO BEN-REQUIRES-AUTH(4)
           MOVE 0.90  TO BEN-ALLOWED-PCT(4).

       1200-READ-CLAIM.
           READ CLAIMS-FILE
               AT END MOVE '10' TO WS-CLM-STATUS
           END-READ.

       2000-PROCESS-CLAIMS.
           ADD 1 TO WS-CLAIMS-PROCESSED
           MOVE SPACES TO WS-DENIAL-REASON
           PERFORM 2100-CHECK-ELIGIBILITY
           IF CLAIM-APPROVED
               PERFORM 2200-CHECK-PROVIDER
           END-IF
           IF CLAIM-APPROVED
               PERFORM 2300-CHECK-AUTHORIZATION
           END-IF
           IF CLAIM-APPROVED
               PERFORM 2400-ADJUDICATE-CLAIM
           END-IF
           PERFORM 2500-GENERATE-EOB
           PERFORM 2600-UPDATE-MEMBER-ACCUMULATORS
           PERFORM 2700-WRITE-OUTPUT
           PERFORM 1200-READ-CLAIM.

       2100-CHECK-ELIGIBILITY.
           MOVE CLM-MEMBER-ID TO MBR-ID
           READ MEMBER-FILE
               INVALID KEY
                   MOVE 'MEMBER NOT FOUND IN SYSTEM' TO WS-DENIAL-REASON
                   MOVE 'D' TO WS-CLAIM-STATUS
           END-READ
           IF MBR-FOUND
               EVALUATE TRUE
                   WHEN MBR-TERMED
                       MOVE 'MEMBER COVERAGE TERMINATED' TO WS-DENIAL-REASON
                       MOVE 'D' TO WS-CLAIM-STATUS
                   WHEN CLM-SERVICE-DATE < MBR-EFFECTIVE-DATE
                       MOVE 'SERVICE BEFORE COVERAGE EFFECTIVE DATE'
                           TO WS-DENIAL-REASON
                       MOVE 'D' TO WS-CLAIM-STATUS
                   WHEN CLM-SERVICE-DATE > MBR-TERM-DATE
                       AND MBR-TERM-DATE NOT = ZEROS
                       MOVE 'SERVICE AFTER COVERAGE TERMINATION DATE'
                           TO WS-DENIAL-REASON
                       MOVE 'D' TO WS-CLAIM-STATUS
                   WHEN OTHER
                       MOVE 'A' TO WS-CLAIM-STATUS
               END-EVALUATE
           END-IF.

       2200-CHECK-PROVIDER.
           MOVE CLM-PROVIDER-NPI TO PRV-NPI
           READ PROVIDER-FILE
               INVALID KEY
                   MOVE 'PROVIDER NOT CREDENTIALED' TO WS-DENIAL-REASON
                   MOVE 'D' TO WS-CLAIM-STATUS
           END-READ.

       2300-CHECK-AUTHORIZATION.
           *> Check if procedure requires prior auth
           PERFORM VARYING BEN-IDX FROM 1 BY 1
               UNTIL BEN-IDX > 10
               IF CLM-PROCEDURE-CODE(1:3) = BEN-PROC-PREFIX(BEN-IDX)
                   IF BEN-REQUIRES-AUTH(BEN-IDX) = 'Y'
                       AND CLM-PRIOR-AUTH = SPACES
                       MOVE 'PRIOR AUTHORIZATION REQUIRED'
                           TO WS-DENIAL-REASON
                       MOVE 'P' TO WS-CLAIM-STATUS
                   END-IF
               END-IF
           END-PERFORM.

       2400-ADJUDICATE-CLAIM.
           *> Calculate allowed amount
           IF PRV-IN-NETWORK
               COMPUTE WS-ALLOWED-AMOUNT ROUNDED =
                   CLM-BILLED-AMOUNT * PRV-CONTRACT-RATE
           ELSE
               *> Out-of-network: use 70% of billed
               COMPUTE WS-ALLOWED-AMOUNT ROUNDED =
                   CLM-BILLED-AMOUNT * 0.70
           END-IF

           *> Apply deductible
           COMPUTE WS-DEDUCTIBLE-APPLY =
               MBR-DEDUCTIBLE-IND - MBR-DEDUCTIBLE-MET
           IF WS-DEDUCTIBLE-APPLY > WS-ALLOWED-AMOUNT
               MOVE WS-ALLOWED-AMOUNT TO WS-DEDUCTIBLE-APPLY
           END-IF
           IF WS-DEDUCTIBLE-APPLY < ZEROS
               MOVE ZEROS TO WS-DEDUCTIBLE-APPLY
           END-IF

           *> Determine copay based on place of service
           EVALUATE TRUE
               WHEN CLM-OFFICE
                   MOVE MBR-COPAY-OFFICE TO WS-COPAY-AMOUNT
               WHEN CLM-EMERGENCY
                   MOVE MBR-COPAY-ER TO WS-COPAY-AMOUNT
               WHEN OTHER
                   MOVE MBR-COPAY-SPEC TO WS-COPAY-AMOUNT
           END-EVALUATE

           *> Coinsurance on amount after deductible
           COMPUTE WS-COINSURANCE-AMT ROUNDED =
               (WS-ALLOWED-AMOUNT - WS-DEDUCTIBLE-APPLY) *
               (1 - MBR-COINSURANCE-PCT / 100)

           *> Plan pays
           COMPUTE WS-PLAN-PAYS =
               WS-ALLOWED-AMOUNT - WS-DEDUCTIBLE-APPLY -
               WS-COINSURANCE-AMT - WS-COPAY-AMOUNT

           IF WS-PLAN-PAYS < ZEROS
               MOVE ZEROS TO WS-PLAN-PAYS
           END-IF

           *> Check OOP max
           IF MBR-OOP-MET >= MBR-OOP-MAX
               MOVE WS-ALLOWED-AMOUNT TO WS-PLAN-PAYS
               MOVE ZEROS TO WS-DEDUCTIBLE-APPLY
               MOVE ZEROS TO WS-COINSURANCE-AMT
               MOVE ZEROS TO WS-COPAY-AMOUNT
           END-IF

           COMPUTE WS-MEMBER-PAYS =
               WS-DEDUCTIBLE-APPLY + WS-COINSURANCE-AMT + WS-COPAY-AMOUNT
           COMPUTE WS-NOT-COVERED =
               CLM-BILLED-AMOUNT - WS-ALLOWED-AMOUNT.

       2500-GENERATE-EOB.
           WRITE EOB-LINE FROM ALL '-'
           MOVE SPACES TO EOB-LINE
           STRING 'CLAIM ID: ' CLM-CLAIM-ID
                  '  MEMBER: ' MBR-NAME
                  '  DATE: ' CLM-SERVICE-DATE
               DELIMITED SIZE INTO EOB-LINE
           WRITE EOB-LINE
           MOVE SPACES TO EOB-LINE
           STRING 'PROVIDER: ' PRV-NAME
                  '  PROCEDURE: ' CLM-PROCEDURE-CODE
                  '  STATUS: ' WS-CLAIM-STATUS
               DELIMITED SIZE INTO EOB-LINE
           WRITE EOB-LINE

           MOVE CLM-BILLED-AMOUNT  TO WF-BILLED
           MOVE WS-ALLOWED-AMOUNT  TO WF-ALLOWED
           MOVE WS-DEDUCTIBLE-APPLY TO WF-DEDUCTIBLE
           MOVE WS-COINSURANCE-AMT TO WF-COINSURANCE
           MOVE WS-COPAY-AMOUNT    TO WF-COPAY
           MOVE WS-PLAN-PAYS       TO WF-PLAN-PAYS
           MOVE WS-MEMBER-PAYS     TO WF-MBR-PAYS

           MOVE SPACES TO EOB-LINE
           STRING 'Billed: ' WF-BILLED
                  '  Allowed: ' WF-ALLOWED
                  '  Not Covered: '
               DELIMITED SIZE INTO EOB-LINE
           WRITE EOB-LINE
           MOVE SPACES TO EOB-LINE
           STRING 'Deductible: ' WF-DEDUCTIBLE
                  '  Coinsurance: ' WF-COINSURANCE
                  '  Copay: ' WF-COPAY
               DELIMITED SIZE INTO EOB-LINE
           WRITE EOB-LINE
           MOVE SPACES TO EOB-LINE
           STRING 'PLAN PAYS: ' WF-PLAN-PAYS
                  '  MEMBER RESPONSIBILITY: ' WF-MBR-PAYS
               DELIMITED SIZE INTO EOB-LINE
           WRITE EOB-LINE
           IF CLAIM-DENIED
               MOVE SPACES TO EOB-LINE
               STRING 'DENIAL REASON: ' WS-DENIAL-REASON
                   DELIMITED SIZE INTO EOB-LINE
               WRITE EOB-LINE
           END-IF.

       2600-UPDATE-MEMBER-ACCUMULATORS.
           IF CLAIM-APPROVED
               ADD WS-DEDUCTIBLE-APPLY TO MBR-DEDUCTIBLE-MET
               ADD WS-MEMBER-PAYS TO MBR-OOP-MET
               REWRITE MEMBER-RECORD
                   INVALID KEY CONTINUE
               END-REWRITE
               ADD WS-ALLOWED-AMOUNT TO WS-TOTAL-ALLOWED
               ADD WS-PLAN-PAYS TO WS-TOTAL-PLAN-PAYS
               ADD WS-MEMBER-PAYS TO WS-TOTAL-MBR-PAYS
           END-IF
           ADD CLM-BILLED-AMOUNT TO WS-TOTAL-BILLED.

       2700-WRITE-OUTPUT.
           EVALUATE TRUE
               WHEN CLAIM-APPROVED
                   ADD 1 TO WS-CLAIMS-APPROVED
                   MOVE SPACES TO PAID-RECORD
                   STRING CLM-CLAIM-ID ' ' CLM-MEMBER-ID ' '
                          WS-PLAN-PAYS ' ' WS-MEMBER-PAYS
                       DELIMITED SIZE INTO PAID-RECORD
                   WRITE PAID-RECORD
               WHEN CLAIM-DENIED
                   ADD 1 TO WS-CLAIMS-DENIED
                   MOVE SPACES TO DENIED-RECORD
                   STRING CLM-CLAIM-ID ' ' CLM-MEMBER-ID ' '
                          WS-DENIAL-REASON
                       DELIMITED SIZE INTO DENIED-RECORD
                   WRITE DENIED-RECORD
               WHEN CLAIM-PENDED
                   ADD 1 TO WS-CLAIMS-PENDED
           END-EVALUATE.

       3000-PRINT-SUMMARY.
           DISPLAY "=== CLAIMS ADJUDICATION SUMMARY ==="
           DISPLAY "Claims Processed : " WS-CLAIMS-PROCESSED
           DISPLAY "Claims Approved  : " WS-CLAIMS-APPROVED
           DISPLAY "Claims Denied    : " WS-CLAIMS-DENIED
           DISPLAY "Claims Pended    : " WS-CLAIMS-PENDED
           DISPLAY "Total Billed     : " WS-TOTAL-BILLED
           DISPLAY "Total Allowed    : " WS-TOTAL-ALLOWED
           DISPLAY "Plan Pays        : " WS-TOTAL-PLAN-PAYS
           DISPLAY "Member Pays      : " WS-TOTAL-MBR-PAYS.

       9000-TERMINATE.
           CLOSE CLAIMS-FILE MEMBER-FILE PROVIDER-FILE
                 EOB-OUTPUT CLAIMS-PAID CLAIMS-DENIED.
