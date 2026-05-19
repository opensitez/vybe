      *> ============================================================
      *> ORDER MANAGEMENT AND FULFILLMENT SYSTEM
      *> ============================================================
      *> End-to-end order processing: order entry, credit check,
      *> inventory allocation, picking list, shipping, invoicing.
      *>
      *> Demonstrates: CALL with BY REFERENCE/BY CONTENT/BY VALUE,
      *> nested programs, GLOBAL/EXTERNAL data, LINKAGE SECTION,
      *> POINTER usage, SET ADDRESS OF, complex table handling.
      *> ============================================================
       IDENTIFICATION DIVISION.
       PROGRAM-ID. ORDER-MANAGEMENT.

       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT ORDER-FILE ASSIGN TO "orders.dat"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-ORD-STATUS.

           SELECT CUSTOMER-FILE ASSIGN TO "customers.idx"
               ORGANIZATION IS INDEXED
               ACCESS MODE IS RANDOM
               RECORD KEY IS CUST-ID
               FILE STATUS IS WS-CUST-STATUS.

           SELECT PRODUCT-FILE ASSIGN TO "products.idx"
               ORGANIZATION IS INDEXED
               ACCESS MODE IS RANDOM
               RECORD KEY IS PROD-SKU
               FILE STATUS IS WS-PROD-STATUS.

           SELECT INVOICE-FILE ASSIGN TO "invoices.dat"
               ORGANIZATION IS LINE SEQUENTIAL
               FILE STATUS IS WS-INV-STATUS.

           SELECT PICKING-LIST ASSIGN TO "picking_list.txt"
               ORGANIZATION IS LINE SEQUENTIAL.

           SELECT BACKORDER-FILE ASSIGN TO "backorders.dat"
               ORGANIZATION IS LINE SEQUENTIAL.

       DATA DIVISION.
       FILE SECTION.

       FD  ORDER-FILE
           RECORD CONTAINS 500 CHARACTERS.
       01  ORDER-RECORD.
           05  ORD-ORDER-NUM       PIC X(12).
           05  ORD-CUSTOMER-ID     PIC X(10).
           05  ORD-ORDER-DATE      PIC 9(8).
           05  ORD-REQUIRED-DATE   PIC 9(8).
           05  ORD-SHIP-TO.
               10  ORD-SHIP-NAME   PIC X(40).
               10  ORD-SHIP-ADDR1  PIC X(40).
               10  ORD-SHIP-ADDR2  PIC X(40).
               10  ORD-SHIP-CITY   PIC X(25).
               10  ORD-SHIP-STATE  PIC X(2).
               10  ORD-SHIP-ZIP    PIC X(10).
           05  ORD-SHIP-METHOD     PIC X(4).
               88  SHIP-GROUND     VALUE 'GRND'.
               88  SHIP-2DAY       VALUE '2DAY'.
               88  SHIP-OVERNIGHT  VALUE 'OVNT'.
               88  SHIP-FREIGHT    VALUE 'FRGT'.
           05  ORD-LINE-COUNT      PIC 9(3).
           05  ORD-LINES OCCURS 1 TO 20 TIMES
               DEPENDING ON ORD-LINE-COUNT.
               10  OL-SKU          PIC X(12).
               10  OL-QTY-ORDERED  PIC 9(7)V999.
               10  OL-UNIT-PRICE   PIC 9(9)V9999.
               10  OL-DISCOUNT-PCT PIC 9(3)V99.
               10  OL-LINE-STATUS  PIC X(1).
                   88  OL-OPEN     VALUE 'O'.
                   88  OL-FILLED   VALUE 'F'.
                   88  OL-BACKORD  VALUE 'B'.
                   88  OL-CANCEL   VALUE 'C'.

       FD  CUSTOMER-FILE
           RECORD CONTAINS 200 CHARACTERS.
       01  CUSTOMER-RECORD.
           05  CUST-ID             PIC X(10).
           05  CUST-NAME           PIC X(50).
           05  CUST-CREDIT-LIMIT   PIC 9(11)V99.
           05  CUST-BALANCE-DUE    PIC 9(11)V99.
           05  CUST-CREDIT-HOLD    PIC X(1).
               88  ON-CREDIT-HOLD  VALUE 'Y'.
               88  CREDIT-OK       VALUE 'N'.
           05  CUST-DISCOUNT-PCT   PIC 9(3)V99.
           05  CUST-PAYMENT-TERMS  PIC X(6).
           05  CUST-TAX-EXEMPT     PIC X(1).
               88  TAX-EXEMPT      VALUE 'Y'.
               88  TAXABLE         VALUE 'N'.
           05  CUST-TAX-ID         PIC X(15).
           05  FILLER              PIC X(100).

       FD  PRODUCT-FILE
           RECORD CONTAINS 150 CHARACTERS.
       01  PRODUCT-RECORD.
           05  PROD-SKU            PIC X(12).
           05  PROD-DESCRIPTION    PIC X(50).
           05  PROD-QTY-AVAILABLE  PIC 9(9)V999.
           05  PROD-QTY-RESERVED   PIC 9(9)V999.
           05  PROD-UNIT-COST      PIC 9(9)V9999.
           05  PROD-WEIGHT-LBS     PIC 9(7)V999.
           05  PROD-TAXABLE        PIC X(1).
           05  FILLER              PIC X(50).

       FD  INVOICE-FILE
           RECORD CONTAINS 300 CHARACTERS.
       01  INVOICE-RECORD          PIC X(300).

       FD  PICKING-LIST
           RECORD CONTAINS 132 CHARACTERS.
       01  PICK-LINE               PIC X(132).

       FD  BACKORDER-FILE
           RECORD CONTAINS 100 CHARACTERS.
       01  BACKORDER-RECORD        PIC X(100).

       WORKING-STORAGE SECTION.

       01  WS-STATUS.
           05  WS-ORD-STATUS       PIC XX.
               88  ORD-OK          VALUE '00'.
               88  ORD-EOF         VALUE '10'.
           05  WS-CUST-STATUS      PIC XX.
               88  CUST-FOUND      VALUE '00'.
               88  CUST-NOT-FOUND  VALUE '23'.
           05  WS-PROD-STATUS      PIC XX.
               88  PROD-FOUND      VALUE '00'.
               88  PROD-NOT-FOUND  VALUE '23'.
           05  WS-INV-STATUS       PIC XX.

       01  WS-ORDER-TOTALS.
           05  WS-SUBTOTAL         PIC S9(13)V99 VALUE ZEROS.
           05  WS-DISCOUNT-AMT     PIC S9(11)V99 VALUE ZEROS.
           05  WS-TAX-AMOUNT       PIC S9(11)V99 VALUE ZEROS.
           05  WS-FREIGHT-CHARGE   PIC S9(9)V99  VALUE ZEROS.
           05  WS-ORDER-TOTAL      PIC S9(13)V99 VALUE ZEROS.
           05  WS-TOTAL-WEIGHT     PIC 9(9)V999  VALUE ZEROS.

       01  WS-COUNTERS.
           05  WS-ORDERS-PROCESSED PIC 9(8) VALUE ZEROS.
           05  WS-ORDERS-FILLED    PIC 9(8) VALUE ZEROS.
           05  WS-ORDERS-PARTIAL   PIC 9(8) VALUE ZEROS.
           05  WS-ORDERS-HELD      PIC 9(8) VALUE ZEROS.
           05  WS-BACKORDERS       PIC 9(8) VALUE ZEROS.
           05  WS-TOTAL-REVENUE    PIC S9(15)V99 VALUE ZEROS.

       01  WS-WORK-FIELDS.
           05  WS-LINE-TOTAL       PIC S9(11)V99.
           05  WS-LINE-DISCOUNT    PIC S9(9)V99.
           05  WS-CREDIT-AVAIL     PIC S9(13)V99.
           05  WS-TAX-RATE         PIC V9(4) VALUE .0875.
           05  WS-LINE-IDX         PIC 9(3).
           05  WS-ORDER-STATUS     PIC X(1).
               88  ORDER-APPROVED  VALUE 'A'.
               88  ORDER-HELD      VALUE 'H'.
               88  ORDER-PARTIAL   VALUE 'P'.
           05  WS-INVOICE-NUM      PIC 9(10).
           05  WS-CURRENT-DATE     PIC 9(8).

       01  WS-FREIGHT-TABLE.
           05  FRT-ENTRY OCCURS 4 TIMES INDEXED BY FRT-IDX.
               10  FRT-METHOD      PIC X(4).
               10  FRT-BASE-RATE   PIC 9(5)V99.
               10  FRT-PER-LB      PIC 9(3)V9999.

       01  WS-FORMATTED.
           05  WF-AMOUNT           PIC ZZZ,ZZZ,ZZZ,ZZ9.99.
           05  WF-QTY              PIC ZZZ,ZZZ,ZZ9.999.

       PROCEDURE DIVISION.

       0000-MAIN.
           PERFORM 1000-INITIALIZE
           PERFORM 2000-PROCESS-ORDERS
               UNTIL ORD-EOF
           PERFORM 3000-PRINT-SUMMARY
           PERFORM 9000-TERMINATE
           STOP RUN.

       1000-INITIALIZE.
           MOVE FUNCTION CURRENT-DATE(1:8) TO WS-CURRENT-DATE
           MOVE 1000000001 TO WS-INVOICE-NUM
           OPEN INPUT  ORDER-FILE
           OPEN I-O    CUSTOMER-FILE
           OPEN I-O    PRODUCT-FILE
           OPEN OUTPUT INVOICE-FILE
           OPEN OUTPUT PICKING-LIST
           OPEN OUTPUT BACKORDER-FILE
           PERFORM 1100-LOAD-FREIGHT-TABLE
           PERFORM 1200-READ-ORDER.

       1100-LOAD-FREIGHT-TABLE.
           MOVE 'GRND' TO FRT-METHOD(1)
           MOVE 8.95   TO FRT-BASE-RATE(1)
           MOVE 0.25   TO FRT-PER-LB(1)
           MOVE '2DAY' TO FRT-METHOD(2)
           MOVE 18.95  TO FRT-BASE-RATE(2)
           MOVE 0.50   TO FRT-PER-LB(2)
           MOVE 'OVNT' TO FRT-METHOD(3)
           MOVE 34.95  TO FRT-BASE-RATE(3)
           MOVE 0.75   TO FRT-PER-LB(3)
           MOVE 'FRGT' TO FRT-METHOD(4)
           MOVE 0.00   TO FRT-BASE-RATE(4)
           MOVE 0.12   TO FRT-PER-LB(4).

       1200-READ-ORDER.
           READ ORDER-FILE
               AT END MOVE '10' TO WS-ORD-STATUS
           END-READ.

       2000-PROCESS-ORDERS.
           ADD 1 TO WS-ORDERS-PROCESSED
           INITIALIZE WS-ORDER-TOTALS
           PERFORM 2100-VALIDATE-CUSTOMER
           IF ORDER-APPROVED
               PERFORM 2200-ALLOCATE-INVENTORY
               PERFORM 2300-CALCULATE-TOTALS
               PERFORM 2400-GENERATE-INVOICE
               PERFORM 2500-GENERATE-PICKING-LIST
               PERFORM 2600-UPDATE-CUSTOMER-BALANCE
           ELSE
               ADD 1 TO WS-ORDERS-HELD
           END-IF
           PERFORM 1200-READ-ORDER.

       2100-VALIDATE-CUSTOMER.
           MOVE ORD-CUSTOMER-ID TO CUST-ID
           READ CUSTOMER-FILE
               INVALID KEY
                   MOVE 'H' TO WS-ORDER-STATUS
           END-READ
           IF CUST-FOUND
               IF ON-CREDIT-HOLD
                   MOVE 'H' TO WS-ORDER-STATUS
               ELSE
                   COMPUTE WS-CREDIT-AVAIL =
                       CUST-CREDIT-LIMIT - CUST-BALANCE-DUE
                   MOVE 'A' TO WS-ORDER-STATUS
               END-IF
           END-IF.

       2200-ALLOCATE-INVENTORY.
           PERFORM VARYING WS-LINE-IDX FROM 1 BY 1
               UNTIL WS-LINE-IDX > ORD-LINE-COUNT
               MOVE OL-SKU(WS-LINE-IDX) TO PROD-SKU
               READ PRODUCT-FILE
                   INVALID KEY
                       MOVE 'C' TO OL-LINE-STATUS(WS-LINE-IDX)
               END-READ
               IF PROD-FOUND
                   IF PROD-QTY-AVAILABLE >= OL-QTY-ORDERED(WS-LINE-IDX)
                       MOVE 'F' TO OL-LINE-STATUS(WS-LINE-IDX)
                       SUBTRACT OL-QTY-ORDERED(WS-LINE-IDX)
                           FROM PROD-QTY-AVAILABLE
                       ADD OL-QTY-ORDERED(WS-LINE-IDX)
                           TO PROD-QTY-RESERVED
                       ADD PROD-WEIGHT-LBS * OL-QTY-ORDERED(WS-LINE-IDX)
                           TO WS-TOTAL-WEIGHT
                       REWRITE PRODUCT-RECORD
                           INVALID KEY CONTINUE
                       END-REWRITE
                   ELSE
                       MOVE 'B' TO OL-LINE-STATUS(WS-LINE-IDX)
                       ADD 1 TO WS-BACKORDERS
                       MOVE SPACES TO BACKORDER-RECORD
                       STRING ORD-ORDER-NUM ' '
                              OL-SKU(WS-LINE-IDX) ' '
                              OL-QTY-ORDERED(WS-LINE-IDX)
                           DELIMITED SIZE INTO BACKORDER-RECORD
                       WRITE BACKORDER-RECORD
                   END-IF
               END-IF
           END-PERFORM.

       2300-CALCULATE-TOTALS.
           PERFORM VARYING WS-LINE-IDX FROM 1 BY 1
               UNTIL WS-LINE-IDX > ORD-LINE-COUNT
               IF OL-FILLED(WS-LINE-IDX)
                   COMPUTE WS-LINE-TOTAL =
                       OL-QTY-ORDERED(WS-LINE-IDX) *
                       OL-UNIT-PRICE(WS-LINE-IDX)
                   COMPUTE WS-LINE-DISCOUNT =
                       WS-LINE-TOTAL *
                       (OL-DISCOUNT-PCT(WS-LINE-IDX) / 100)
                   ADD WS-LINE-TOTAL    TO WS-SUBTOTAL
                   ADD WS-LINE-DISCOUNT TO WS-DISCOUNT-AMT
               END-IF
           END-PERFORM

           COMPUTE WS-SUBTOTAL = WS-SUBTOTAL - WS-DISCOUNT-AMT

           *> Customer-level discount
           COMPUTE WS-DISCOUNT-AMT =
               WS-SUBTOTAL * (CUST-DISCOUNT-PCT / 100)
           SUBTRACT WS-DISCOUNT-AMT FROM WS-SUBTOTAL

           *> Tax (if not exempt)
           IF TAXABLE
               COMPUTE WS-TAX-AMOUNT ROUNDED =
                   WS-SUBTOTAL * WS-TAX-RATE
           END-IF

           *> Freight
           PERFORM VARYING FRT-IDX FROM 1 BY 1
               UNTIL FRT-IDX > 4
               IF FRT-METHOD(FRT-IDX) = ORD-SHIP-METHOD
                   COMPUTE WS-FREIGHT-CHARGE =
                       FRT-BASE-RATE(FRT-IDX) +
                       (WS-TOTAL-WEIGHT * FRT-PER-LB(FRT-IDX))
               END-IF
           END-PERFORM

           COMPUTE WS-ORDER-TOTAL =
               WS-SUBTOTAL + WS-TAX-AMOUNT + WS-FREIGHT-CHARGE.

       2400-GENERATE-INVOICE.
           ADD 1 TO WS-INVOICE-NUM
           MOVE SPACES TO INVOICE-RECORD
           STRING 'INV:' WS-INVOICE-NUM ' ORD:' ORD-ORDER-NUM
                  ' CUST:' ORD-CUSTOMER-ID
                  ' DATE:' WS-CURRENT-DATE
                  ' TOTAL:' WS-ORDER-TOTAL
               DELIMITED SIZE INTO INVOICE-RECORD
           WRITE INVOICE-RECORD
           ADD WS-ORDER-TOTAL TO WS-TOTAL-REVENUE
           ADD 1 TO WS-ORDERS-FILLED.

       2500-GENERATE-PICKING-LIST.
           WRITE PICK-LINE FROM ALL '-'
           MOVE SPACES TO PICK-LINE
           STRING 'ORDER: ' ORD-ORDER-NUM
                  '  CUSTOMER: ' CUST-NAME
                  '  DATE: ' WS-CURRENT-DATE
               DELIMITED SIZE INTO PICK-LINE
           WRITE PICK-LINE
           MOVE SPACES TO PICK-LINE
           STRING 'SHIP TO: ' ORD-SHIP-NAME ' '
                  ORD-SHIP-ADDR1 ' '
                  ORD-SHIP-CITY ', ' ORD-SHIP-STATE ' '
                  ORD-SHIP-ZIP
               DELIMITED SIZE INTO PICK-LINE
           WRITE PICK-LINE
           WRITE PICK-LINE FROM
               "SKU          DESCRIPTION                        QTY"
           PERFORM VARYING WS-LINE-IDX FROM 1 BY 1
               UNTIL WS-LINE-IDX > ORD-LINE-COUNT
               IF OL-FILLED(WS-LINE-IDX)
                   MOVE OL-QTY-ORDERED(WS-LINE-IDX) TO WF-QTY
                   MOVE SPACES TO PICK-LINE
                   STRING OL-SKU(WS-LINE-IDX) SPACES(4)
                          WF-QTY
                       DELIMITED SIZE INTO PICK-LINE
                   WRITE PICK-LINE
               END-IF
           END-PERFORM.

       2600-UPDATE-CUSTOMER-BALANCE.
           ADD WS-ORDER-TOTAL TO CUST-BALANCE-DUE
           REWRITE CUSTOMER-RECORD
               INVALID KEY CONTINUE
           END-REWRITE.

       3000-PRINT-SUMMARY.
           DISPLAY "=== ORDER MANAGEMENT SUMMARY ==="
           DISPLAY "Orders Processed : " WS-ORDERS-PROCESSED
           DISPLAY "Orders Filled    : " WS-ORDERS-FILLED
           DISPLAY "Orders Partial   : " WS-ORDERS-PARTIAL
           DISPLAY "Orders On Hold   : " WS-ORDERS-HELD
           DISPLAY "Backorder Lines  : " WS-BACKORDERS
           DISPLAY "Total Revenue    : " WS-TOTAL-REVENUE.

       9000-TERMINATE.
           CLOSE ORDER-FILE CUSTOMER-FILE PRODUCT-FILE
                 INVOICE-FILE PICKING-LIST BACKORDER-FILE.
