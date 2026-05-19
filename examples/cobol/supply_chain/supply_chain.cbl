      *> ============================================================
      *> SUPPLY CHAIN MANAGEMENT SYSTEM
      *> ============================================================
      *> End-to-end supply chain: demand forecasting, purchase order
      *> generation, supplier management, receiving, quality control,
      *> landed cost calculation, supplier scorecards.
      *>
      *> Demonstrates: COBOL 2014 VALIDATE, complex multi-file
      *> processing, SORT with USING/GIVING, table-driven logic,
      *> PERFORM with inline UNTIL, COMPUTE with multiple operands,
      *> nested CALL, EXTERNAL data items.
      *> ============================================================
       IDENTIFICATION DIVISION.
       PROGRAM-ID. SUPPLY-CHAIN.

       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT SUPPLIER-MASTER ASSIGN TO "suppliers.idx"
               ORGANIZATION IS INDEXED
               ACCESS MODE IS DYNAMIC
               RECORD KEY IS SUP-ID
               FILE STATUS IS WS-SUP-STATUS.

           SELECT DEMAND-HISTORY ASSIGN TO "demand_history.dat"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-DEM-STATUS.

           SELECT OPEN-PO-FILE ASSIGN TO "open_purchase_orders.idx"
               ORGANIZATION IS INDEXED
               ACCESS MODE IS DYNAMIC
               RECORD KEY IS PO-NUMBER
               ALTERNATE RECORD KEY IS PO-SUPPLIER-ITEM
                   WITH DUPLICATES
               FILE STATUS IS WS-PO-STATUS.

           SELECT RECEIVING-FILE ASSIGN TO "receiving_log.dat"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-RCV-STATUS.

           SELECT ITEM-MASTER ASSIGN TO "items.idx"
               ORGANIZATION IS INDEXED
               ACCESS MODE IS DYNAMIC
               RECORD KEY IS ITEM-CODE
               FILE STATUS IS WS-ITM-STATUS.

           SELECT PO-GENERATION ASSIGN TO "new_purchase_orders.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT SUPPLIER-SCORECARD ASSIGN TO "supplier_scorecard.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT FORECAST-OUTPUT ASSIGN TO "demand_forecast.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

       DATA DIVISION.
       FILE SECTION.

       FD  SUPPLIER-MASTER
           RECORD CONTAINS 300 CHARACTERS.
       01  SUPPLIER-RECORD.
           05  SUP-ID              PIC X(8).
           05  SUP-NAME            PIC X(50).
           05  SUP-CONTACT         PIC X(40).
           05  SUP-EMAIL           PIC X(60).
           05  SUP-PHONE           PIC X(15).
           05  SUP-COUNTRY         PIC X(3).
           05  SUP-CURRENCY        PIC X(3).
           05  SUP-PAYMENT-TERMS   PIC X(6).
           05  SUP-LEAD-TIME-DAYS  PIC 9(3).
           05  SUP-MIN-ORDER-AMT   PIC 9(9)V99 COMP-3.
           05  SUP-PERFORMANCE.
               10  SUP-ON-TIME-PCT     PIC 9(3)V99 COMP-3.
               10  SUP-QUALITY-PCT     PIC 9(3)V99 COMP-3.
               10  SUP-FILL-RATE-PCT   PIC 9(3)V99 COMP-3.
               10  SUP-TOTAL-ORDERS    PIC 9(8)    COMP.
               10  SUP-LATE-ORDERS     PIC 9(8)    COMP.
               10  SUP-REJECTED-LOTS   PIC 9(6)    COMP.
               10  SUP-TOTAL-SPEND     PIC 9(13)V99 COMP-3.
               10  SUP-SCORE           PIC 9(3)V99 COMP-3.
           05  SUP-ACTIVE-FLAG     PIC X(1).
               88  SUP-ACTIVE      VALUE 'Y'.
               88  SUP-INACTIVE    VALUE 'N'.
           05  FILLER              PIC X(50).

       FD  DEMAND-HISTORY
           RECORD CONTAINS 80 CHARACTERS.
       01  DEMAND-RECORD.
           05  DEM-ITEM-CODE       PIC X(12).
           05  DEM-PERIOD          PIC 9(6).
           05  DEM-QUANTITY        PIC 9(9)V999.
           05  DEM-UNIT-COST       PIC 9(9)V9999.
           05  FILLER              PIC X(41).

       FD  OPEN-PO-FILE
           RECORD CONTAINS 300 CHARACTERS.
       01  PO-RECORD.
           05  PO-NUMBER           PIC X(12).
           05  PO-SUPPLIER-ITEM.
               10  PO-SUPPLIER-ID  PIC X(8).
               10  PO-ITEM-CODE    PIC X(12).
           05  PO-DATE-ISSUED      PIC 9(8).
           05  PO-DATE-EXPECTED    PIC 9(8).
           05  PO-DATE-RECEIVED    PIC 9(8).
           05  PO-STATUS           PIC X(2).
               88  PO-OPEN         VALUE 'OP'.
               88  PO-PARTIAL      VALUE 'PA'.
               88  PO-RECEIVED     VALUE 'RC'.
               88  PO-CANCELLED    VALUE 'CA'.
               88  PO-OVERDUE      VALUE 'OD'.
           05  PO-QTY-ORDERED      PIC 9(9)V999 COMP-3.
           05  PO-QTY-RECEIVED     PIC 9(9)V999 COMP-3.
           05  PO-UNIT-PRICE       PIC 9(9)V9999 COMP-3.
           05  PO-FREIGHT-COST     PIC 9(7)V99   COMP-3.
           05  PO-DUTY-RATE        PIC V9(4)     COMP-3.
           05  PO-LANDED-COST      PIC 9(9)V9999 COMP-3.
           05  PO-QUALITY-RESULT   PIC X(1).
               88  QC-PASSED       VALUE 'P'.
               88  QC-FAILED       VALUE 'F'.
               88  QC-PENDING      VALUE ' '.
           05  PO-REJECTION-REASON PIC X(40).
           05  FILLER              PIC X(100).

       FD  RECEIVING-FILE
           RECORD CONTAINS 100 CHARACTERS.
       01  RECEIVING-RECORD.
           05  RCV-PO-NUMBER       PIC X(12).
           05  RCV-DATE            PIC 9(8).
           05  RCV-QTY-RECEIVED    PIC 9(9)V999.
           05  RCV-QTY-REJECTED    PIC 9(9)V999.
           05  RCV-REJECTION-CODE  PIC X(4).
           05  FILLER              PIC X(59).

       FD  ITEM-MASTER
           RECORD CONTAINS 250 CHARACTERS.
       01  ITEM-RECORD.
           05  ITEM-CODE           PIC X(12).
           05  ITEM-DESCRIPTION    PIC X(50).
           05  ITEM-CATEGORY       PIC X(6).
           05  ITEM-UOM            PIC X(4).
           05  ITEM-PREFERRED-SUP  PIC X(8).
           05  ITEM-BACKUP-SUP     PIC X(8).
           05  ITEM-QTY-ON-HAND    PIC 9(9)V999 COMP-3.
           05  ITEM-QTY-ON-ORDER   PIC 9(9)V999 COMP-3.
           05  ITEM-SAFETY-STOCK   PIC 9(9)V999 COMP-3.
           05  ITEM-REORDER-POINT  PIC 9(9)V999 COMP-3.
           05  ITEM-EOQ            PIC 9(9)V999 COMP-3.
           05  ITEM-UNIT-COST      PIC 9(9)V9999 COMP-3.
           05  ITEM-DEMAND-HISTORY OCCURS 12 TIMES.
               10  HIST-PERIOD     PIC 9(6).
               10  HIST-QTY        PIC 9(9)V999 COMP-3.
           05  ITEM-FORECAST-QTY   PIC 9(9)V999 COMP-3.
           05  ITEM-FORECAST-METHOD PIC X(3).
           05  FILLER              PIC X(50).

       FD  PO-GENERATION
           RECORD CONTAINS 200 CHARACTERS.
       01  PO-GEN-LINE             PIC X(200).

       FD  SUPPLIER-SCORECARD
           RECORD CONTAINS 132 CHARACTERS.
       01  SCORE-LINE              PIC X(132).

       FD  FORECAST-OUTPUT
           RECORD CONTAINS 132 CHARACTERS.
       01  FCST-LINE               PIC X(132).

       WORKING-STORAGE SECTION.

       01  WS-STATUS.
           05  WS-SUP-STATUS       PIC XX.
               88  SUP-OK          VALUE '00'.
               88  SUP-NOT-FOUND   VALUE '23'.
               88  SUP-EOF         VALUE '10'.
           05  WS-DEM-STATUS       PIC XX.
               88  DEM-OK          VALUE '00'.
               88  DEM-EOF         VALUE '10'.
           05  WS-PO-STATUS        PIC XX.
               88  PO-OK           VALUE '00'.
               88  PO-NOT-FOUND    VALUE '23'.
               88  PO-EOF          VALUE '10'.
           05  WS-RCV-STATUS       PIC XX.
               88  RCV-OK          VALUE '00'.
               88  RCV-EOF         VALUE '10'.
           05  WS-ITM-STATUS       PIC XX.
               88  ITM-OK          VALUE '00'.
               88  ITM-NOT-FOUND   VALUE '23'.
               88  ITM-EOF         VALUE '10'.

       01  WS-FORECAST-WORK.
           05  WS-ALPHA            PIC V9(4) VALUE .3000.
           05  WS-FORECAST-NEXT    PIC 9(9)V999 COMP-3.
           05  WS-SMOOTHED         PIC 9(9)V999 COMP-3.
           05  WS-TOTAL-DEMAND     PIC 9(11)V999 COMP-3.
           05  WS-AVG-DEMAND       PIC 9(9)V999 COMP-3.
           05  WS-HIST-IDX         PIC 9(2).
           05  WS-PERIODS-WITH-DATA PIC 9(2).

       01  WS-EOQ-WORK.
           05  WS-ANNUAL-DEMAND    PIC 9(11)V999 COMP-3.
           05  WS-ORDERING-COST    PIC 9(7)V99   COMP-3 VALUE 150.00.
           05  WS-HOLDING-RATE     PIC V9(4)     COMP-3 VALUE .2500.
           05  WS-EOQ-CALC         PIC 9(9)V999  COMP-3.

       01  WS-LANDED-COST-WORK.
           05  WS-PRODUCT-COST     PIC 9(11)V99 COMP-3.
           05  WS-FREIGHT-COST     PIC 9(9)V99  COMP-3.
           05  WS-DUTY-AMOUNT      PIC 9(9)V99  COMP-3.
           05  WS-INSURANCE        PIC 9(7)V99  COMP-3.
           05  WS-TOTAL-LANDED     PIC 9(11)V99 COMP-3.
           05  WS-LANDED-PER-UNIT  PIC 9(9)V9999 COMP-3.

       01  WS-COUNTERS.
           05  WS-ITEMS-ANALYZED   PIC 9(8) VALUE ZEROS.
           05  WS-POS-GENERATED    PIC 9(6) VALUE ZEROS.
           05  WS-RECEIPTS-PROC    PIC 9(8) VALUE ZEROS.
           05  WS-QC-FAILURES      PIC 9(6) VALUE ZEROS.
           05  WS-OVERDUE-POS      PIC 9(6) VALUE ZEROS.
           05  WS-TOTAL-PO-VALUE   PIC 9(15)V99 VALUE ZEROS.

       01  WS-CURRENT-DATE         PIC 9(8).
       01  WS-CURRENT-PERIOD       PIC 9(6).
       01  WS-NEXT-PO-NUM          PIC 9(12) VALUE 100000000001.

       01  WS-FORMATTED.
           05  WF-AMOUNT           PIC ZZZ,ZZZ,ZZZ,ZZ9.99.
           05  WF-QTY              PIC ZZZ,ZZZ,ZZ9.999.
           05  WF-PCT              PIC ZZZ9.99.
           05  WF-SCORE            PIC ZZ9.99.

       PROCEDURE DIVISION.

       0000-MAIN.
           PERFORM 1000-INITIALIZE
           PERFORM 2000-PROCESS-RECEIPTS
               UNTIL RCV-EOF
           PERFORM 3000-DEMAND-FORECASTING
           PERFORM 4000-GENERATE-PURCHASE-ORDERS
           PERFORM 5000-SUPPLIER-SCORECARD
           PERFORM 6000-PRINT-SUMMARY
           PERFORM 9000-TERMINATE
           STOP RUN.

       1000-INITIALIZE.
           MOVE FUNCTION CURRENT-DATE(1:8) TO WS-CURRENT-DATE
           MOVE WS-CURRENT-DATE(1:6) TO WS-CURRENT-PERIOD
           OPEN I-O    SUPPLIER-MASTER
           OPEN INPUT  DEMAND-HISTORY
           OPEN I-O    OPEN-PO-FILE
           OPEN INPUT  RECEIVING-FILE
           OPEN I-O    ITEM-MASTER
           OPEN OUTPUT PO-GENERATION
           OPEN OUTPUT SUPPLIER-SCORECARD
           OPEN OUTPUT FORECAST-OUTPUT
           PERFORM 1100-READ-RECEIVING.

       1100-READ-RECEIVING.
           READ RECEIVING-FILE
               AT END MOVE '10' TO WS-RCV-STATUS
           END-READ.

       2000-PROCESS-RECEIPTS.
           MOVE RCV-PO-NUMBER TO PO-NUMBER
           READ OPEN-PO-FILE
               INVALID KEY
                   PERFORM 1100-READ-RECEIVING
                   STOP RUN
           END-READ
           IF PO-OK
               ADD RCV-QTY-RECEIVED TO PO-QTY-RECEIVED
               MOVE RCV-DATE TO PO-DATE-RECEIVED

               *> Calculate landed cost
               PERFORM 2100-CALC-LANDED-COST

               *> Quality control check
               IF RCV-QTY-REJECTED > ZEROS
                   MOVE 'F' TO PO-QUALITY-RESULT
                   ADD 1 TO WS-QC-FAILURES
                   MOVE SUP-ID TO PO-SUPPLIER-ID
                   MOVE PO-SUPPLIER-ID TO SUP-ID
                   READ SUPPLIER-MASTER
                       INVALID KEY CONTINUE
                   END-READ
                   IF SUP-OK
                       ADD 1 TO SUP-REJECTED-LOTS
                       REWRITE SUPPLIER-RECORD
                           INVALID KEY CONTINUE
                       END-REWRITE
                   END-IF
               ELSE
                   MOVE 'P' TO PO-QUALITY-RESULT
               END-IF

               *> Check if fully received
               IF PO-QTY-RECEIVED >= PO-QTY-ORDERED
                   MOVE 'RC' TO PO-STATUS
               ELSE
                   MOVE 'PA' TO PO-STATUS
               END-IF

               *> Update on-time delivery for supplier
               PERFORM 2200-UPDATE-SUPPLIER-PERFORMANCE

               REWRITE PO-RECORD
                   INVALID KEY CONTINUE
               END-REWRITE

               *> Update item on-hand
               MOVE PO-ITEM-CODE TO ITEM-CODE
               READ ITEM-MASTER
                   INVALID KEY CONTINUE
               END-READ
               IF ITM-OK
                   ADD RCV-QTY-RECEIVED TO ITEM-QTY-ON-HAND
                   SUBTRACT RCV-QTY-RECEIVED FROM ITEM-QTY-ON-ORDER
                   REWRITE ITEM-RECORD
                       INVALID KEY CONTINUE
                   END-REWRITE
               END-IF
               ADD 1 TO WS-RECEIPTS-PROC
           END-IF
           PERFORM 1100-READ-RECEIVING.

       2100-CALC-LANDED-COST.
           COMPUTE WS-PRODUCT-COST =
               PO-QTY-ORDERED * PO-UNIT-PRICE
           MOVE PO-FREIGHT-COST TO WS-FREIGHT-COST
           COMPUTE WS-DUTY-AMOUNT ROUNDED =
               WS-PRODUCT-COST * PO-DUTY-RATE
           COMPUTE WS-INSURANCE ROUNDED =
               WS-PRODUCT-COST * 0.005
           COMPUTE WS-TOTAL-LANDED =
               WS-PRODUCT-COST + WS-FREIGHT-COST +
               WS-DUTY-AMOUNT + WS-INSURANCE
           IF PO-QTY-ORDERED > ZEROS
               COMPUTE PO-LANDED-COST ROUNDED =
                   WS-TOTAL-LANDED / PO-QTY-ORDERED
           END-IF.

       2200-UPDATE-SUPPLIER-PERFORMANCE.
           MOVE PO-SUPPLIER-ID TO SUP-ID
           READ SUPPLIER-MASTER
               INVALID KEY CONTINUE
           END-READ
           IF SUP-OK
               ADD 1 TO SUP-TOTAL-ORDERS
               *> Check on-time delivery
               IF PO-DATE-RECEIVED > PO-DATE-EXPECTED
                   ADD 1 TO SUP-LATE-ORDERS
               END-IF
               *> Recalculate on-time percentage
               IF SUP-TOTAL-ORDERS > ZEROS
                   COMPUTE SUP-ON-TIME-PCT ROUNDED =
                       ((SUP-TOTAL-ORDERS - SUP-LATE-ORDERS) * 100) /
                       SUP-TOTAL-ORDERS
               END-IF
               ADD WS-PRODUCT-COST TO SUP-TOTAL-SPEND
               *> Composite score: 40% on-time, 40% quality, 20% fill rate
               COMPUTE SUP-SCORE ROUNDED =
                   (SUP-ON-TIME-PCT * 0.40) +
                   (SUP-QUALITY-PCT * 0.40) +
                   (SUP-FILL-RATE-PCT * 0.20)
               REWRITE SUPPLIER-RECORD
                   INVALID KEY CONTINUE
               END-REWRITE
           END-IF.

       3000-DEMAND-FORECASTING.
           WRITE FCST-LINE FROM "DEMAND FORECAST REPORT"
           WRITE FCST-LINE FROM ALL '='
           WRITE FCST-LINE FROM
               "ITEM CODE    DESCRIPTION                    " &
               "AVG DEMAND   FORECAST     METHOD   EOQ"
           WRITE FCST-LINE FROM ALL '-'

           MOVE LOW-VALUES TO ITEM-CODE
           START ITEM-MASTER KEY >= ITEM-CODE
               INVALID KEY STOP RUN
           END-START
           PERFORM 3100-FORECAST-SCAN UNTIL ITM-EOF.

       3100-FORECAST-SCAN.
           READ ITEM-MASTER NEXT
               AT END MOVE '10' TO WS-ITM-STATUS
           END-READ
           IF NOT ITM-EOF
               ADD 1 TO WS-ITEMS-ANALYZED
               PERFORM 3200-CALC-FORECAST
               PERFORM 3300-CALC-EOQ
               MOVE ITEM-FORECAST-QTY TO WF-QTY
               MOVE ITEM-EOQ TO WF-AMOUNT
               MOVE SPACES TO FCST-LINE
               STRING ITEM-CODE ' '
                      ITEM-DESCRIPTION ' '
                      WS-AVG-DEMAND ' '
                      WF-QTY ' '
                      ITEM-FORECAST-METHOD ' '
                      WF-AMOUNT
                   DELIMITED SIZE INTO FCST-LINE
               WRITE FCST-LINE
               REWRITE ITEM-RECORD
                   INVALID KEY CONTINUE
               END-REWRITE
           END-IF.

       3200-CALC-FORECAST.
           *> Exponential smoothing forecast
           MOVE ZEROS TO WS-TOTAL-DEMAND WS-PERIODS-WITH-DATA
           MOVE ITEM-HIST-QTY(1) TO WS-SMOOTHED

           PERFORM VARYING WS-HIST-IDX FROM 1 BY 1
               UNTIL WS-HIST-IDX > 12
               IF ITEM-HIST-QTY(WS-HIST-IDX) > ZEROS
                   ADD ITEM-HIST-QTY(WS-HIST-IDX) TO WS-TOTAL-DEMAND
                   ADD 1 TO WS-PERIODS-WITH-DATA
                   *> Exponential smoothing: F(t+1) = alpha*A(t) + (1-alpha)*F(t)
                   COMPUTE WS-SMOOTHED ROUNDED =
                       WS-ALPHA * ITEM-HIST-QTY(WS-HIST-IDX) +
                       (1 - WS-ALPHA) * WS-SMOOTHED
               END-IF
           END-PERFORM

           IF WS-PERIODS-WITH-DATA > ZEROS
               COMPUTE WS-AVG-DEMAND ROUNDED =
                   WS-TOTAL-DEMAND / WS-PERIODS-WITH-DATA
           END-IF

           MOVE WS-SMOOTHED TO ITEM-FORECAST-QTY
           MOVE 'EXP' TO ITEM-FORECAST-METHOD.

       3300-CALC-EOQ.
           *> Economic Order Quantity: sqrt(2*D*S / H*C)
           COMPUTE WS-ANNUAL-DEMAND = ITEM-FORECAST-QTY * 12
           IF ITEM-UNIT-COST > ZEROS AND WS-ANNUAL-DEMAND > ZEROS
               COMPUTE WS-EOQ-CALC ROUNDED =
                   FUNCTION SQRT(
                       2 * WS-ANNUAL-DEMAND * WS-ORDERING-COST /
                       (WS-HOLDING-RATE * ITEM-UNIT-COST))
               MOVE WS-EOQ-CALC TO ITEM-EOQ
           END-IF.

       4000-GENERATE-PURCHASE-ORDERS.
           WRITE PO-GEN-LINE FROM "PURCHASE ORDER GENERATION"
           WRITE PO-GEN-LINE FROM ALL '='

           MOVE LOW-VALUES TO ITEM-CODE
           START ITEM-MASTER KEY >= ITEM-CODE
               INVALID KEY STOP RUN
           END-START
           PERFORM 4100-PO-GENERATION-SCAN UNTIL ITM-EOF.

       4100-PO-GENERATION-SCAN.
           READ ITEM-MASTER NEXT
               AT END MOVE '10' TO WS-ITM-STATUS
           END-READ
           IF NOT ITM-EOF
               *> Check if reorder needed
               COMPUTE WS-TOTAL-DEMAND =
                   ITEM-QTY-ON-HAND + ITEM-QTY-ON-ORDER
               IF WS-TOTAL-DEMAND <= ITEM-REORDER-POINT
                   PERFORM 4200-CREATE-PO
               END-IF
           END-IF.

       4200-CREATE-PO.
           ADD 1 TO WS-NEXT-PO-NUM
           ADD 1 TO WS-POS-GENERATED
           COMPUTE WS-TOTAL-LANDED =
               ITEM-EOQ * ITEM-UNIT-COST
           ADD WS-TOTAL-LANDED TO WS-TOTAL-PO-VALUE

           MOVE SPACES TO PO-GEN-LINE
           STRING 'PO#' WS-NEXT-PO-NUM
                  ' ITEM:' ITEM-CODE
                  ' SUP:' ITEM-PREFERRED-SUP
                  ' QTY:' ITEM-EOQ
                  ' EST-VALUE:' WS-TOTAL-LANDED
               DELIMITED SIZE INTO PO-GEN-LINE
           WRITE PO-GEN-LINE.

       5000-SUPPLIER-SCORECARD.
           WRITE SCORE-LINE FROM "SUPPLIER PERFORMANCE SCORECARD"
           WRITE SCORE-LINE FROM ALL '='
           WRITE SCORE-LINE FROM
               "SUPPLIER  NAME                    ON-TIME%  " &
               "QUALITY%  FILL-RATE%  SCORE   TOTAL-SPEND"
           WRITE SCORE-LINE FROM ALL '-'

           MOVE LOW-VALUES TO SUP-ID
           START SUPPLIER-MASTER KEY >= SUP-ID
               INVALID KEY STOP RUN
           END-START
           PERFORM 5100-SCORECARD-SCAN UNTIL SUP-EOF.

       5100-SCORECARD-SCAN.
           READ SUPPLIER-MASTER NEXT
               AT END MOVE '10' TO WS-SUP-STATUS
           END-READ
           IF NOT SUP-EOF AND SUP-ACTIVE
               MOVE SUP-ON-TIME-PCT   TO WF-PCT
               MOVE SUP-SCORE         TO WF-SCORE
               MOVE SUP-TOTAL-SPEND   TO WF-AMOUNT
               MOVE SPACES TO SCORE-LINE
               STRING SUP-ID ' '
                      SUP-NAME ' '
                      WF-PCT ' '
                      SUP-QUALITY-PCT ' '
                      SUP-FILL-RATE-PCT ' '
                      WF-SCORE ' '
                      WF-AMOUNT
                   DELIMITED SIZE INTO SCORE-LINE
               WRITE SCORE-LINE
           END-IF.

       6000-PRINT-SUMMARY.
           DISPLAY "=== SUPPLY CHAIN SUMMARY ==="
           DISPLAY "Items Analyzed     : " WS-ITEMS-ANALYZED
           DISPLAY "Receipts Processed : " WS-RECEIPTS-PROC
           DISPLAY "QC Failures        : " WS-QC-FAILURES
           DISPLAY "POs Generated      : " WS-POS-GENERATED
           DISPLAY "Total PO Value     : " WS-TOTAL-PO-VALUE
           DISPLAY "Overdue POs        : " WS-OVERDUE-POS.

       9000-TERMINATE.
           CLOSE SUPPLIER-MASTER DEMAND-HISTORY OPEN-PO-FILE
                 RECEIVING-FILE ITEM-MASTER
                 PO-GENERATION SUPPLIER-SCORECARD FORECAST-OUTPUT.
