use vybec::parser_cobol::parse;
use vybec::compiler_cobol::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

fn p(data: &str, body: &str) -> String {
    format!("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.", data, body)
}

fn d() -> &'static str {
    "01 WS-ID    PIC 9(10) VALUE 0.\n01 WS-NAME  PIC X(50).\n01 WS-AMT   PIC 9(10)V99 VALUE 0.\n01 WS-DSN   PIC X(100) VALUE \"sqlite:test.db\".\n01 SQLCODE  PIC S9(9) VALUE 0."
}

// ═══════════════════════════════════════════════════════════
// CONNECT
// ═══════════════════════════════════════════════════════════
#[test]
fn sql_connect() {
    compile_ok(&p(d(), "    EXEC SQL CONNECT :WS-DSN END-EXEC."));
}

// ═══════════════════════════════════════════════════════════
// SELECT INTO (single row)
// ═══════════════════════════════════════════════════════════
#[test]
fn sql_select_into() {
    compile_ok(&p(d(),
        "    EXEC SQL\n        SELECT NAME, BALANCE\n        INTO :WS-NAME, :WS-AMT\n        FROM CUSTOMERS\n        WHERE ID = :WS-ID\n    END-EXEC."));
}

#[test]
fn sql_select_simple() {
    compile_ok(&p(d(),
        "    EXEC SQL\n        SELECT NAME INTO :WS-NAME FROM USERS WHERE ID = 1\n    END-EXEC."));
}

#[test]
fn sql_select_check_sqlcode() {
    compile_ok(&p(d(),
        "    EXEC SQL\n        SELECT NAME INTO :WS-NAME FROM USERS WHERE ID = :WS-ID\n    END-EXEC.\n    IF SQLCODE = 0\n        DISPLAY WS-NAME\n    ELSE\n        DISPLAY \"Not found\"\n    END-IF."));
}

// ═══════════════════════════════════════════════════════════
// INSERT
// ═══════════════════════════════════════════════════════════
#[test]
fn sql_insert() {
    compile_ok(&p(d(),
        "    EXEC SQL\n        INSERT INTO CUSTOMERS (ID, NAME, BALANCE)\n        VALUES (:WS-ID, :WS-NAME, :WS-AMT)\n    END-EXEC."));
}

#[test]
fn sql_insert_simple() {
    compile_ok(&p(d(),
        "    EXEC SQL INSERT INTO LOG (MSG) VALUES ('Hello') END-EXEC."));
}

// ═══════════════════════════════════════════════════════════
// UPDATE
// ═══════════════════════════════════════════════════════════
#[test]
fn sql_update() {
    compile_ok(&p(d(),
        "    EXEC SQL\n        UPDATE CUSTOMERS\n        SET NAME = :WS-NAME, BALANCE = :WS-AMT\n        WHERE ID = :WS-ID\n    END-EXEC."));
}

// ═══════════════════════════════════════════════════════════
// DELETE
// ═══════════════════════════════════════════════════════════
#[test]
fn sql_delete() {
    compile_ok(&p(d(),
        "    EXEC SQL\n        DELETE FROM CUSTOMERS WHERE ID = :WS-ID\n    END-EXEC."));
}

// ═══════════════════════════════════════════════════════════
// COMMIT / ROLLBACK
// ═══════════════════════════════════════════════════════════
#[test]
fn sql_commit() {
    compile_ok(&p(d(), "    EXEC SQL COMMIT END-EXEC."));
}

#[test]
fn sql_rollback() {
    compile_ok(&p(d(), "    EXEC SQL ROLLBACK END-EXEC."));
}

// ═══════════════════════════════════════════════════════════
// CURSOR (multi-row processing)
// ═══════════════════════════════════════════════════════════
#[test]
fn sql_declare_cursor() {
    compile_ok(&p(d(),
        "    EXEC SQL\n        DECLARE C1 CURSOR FOR\n        SELECT ID, NAME FROM CUSTOMERS\n    END-EXEC."));
}

#[test]
fn sql_open_cursor() {
    compile_ok(&p(d(), "    EXEC SQL OPEN C1 END-EXEC."));
}

#[test]
fn sql_fetch_cursor() {
    compile_ok(&p(d(),
        "    EXEC SQL FETCH C1 INTO :WS-ID, :WS-NAME END-EXEC."));
}

#[test]
fn sql_close_cursor() {
    compile_ok(&p(d(), "    EXEC SQL CLOSE C1 END-EXEC."));
}

// ═══════════════════════════════════════════════════════════
// FULL CURSOR LOOP PATTERN
// ═══════════════════════════════════════════════════════════
#[test]
fn sql_cursor_loop() {
    compile_ok(&p(d(), r#"
    EXEC SQL CONNECT :WS-DSN END-EXEC.
    EXEC SQL
        DECLARE CUST-CURSOR CURSOR FOR
        SELECT ID, NAME, BALANCE FROM CUSTOMERS
    END-EXEC.
    EXEC SQL OPEN CUST-CURSOR END-EXEC.
    PERFORM UNTIL SQLCODE NOT = 0
        EXEC SQL
            FETCH CUST-CURSOR INTO :WS-ID, :WS-NAME, :WS-AMT
        END-EXEC
        IF SQLCODE = 0
            DISPLAY "ID: " WS-ID " Name: " WS-NAME
        END-IF
    END-PERFORM.
    EXEC SQL CLOSE CUST-CURSOR END-EXEC.
"#));
}

// ═══════════════════════════════════════════════════════════
// TRANSACTION PATTERN
// ═══════════════════════════════════════════════════════════
#[test]
fn sql_transaction() {
    compile_ok(&p(d(), r#"
    EXEC SQL CONNECT :WS-DSN END-EXEC.
    EXEC SQL
        INSERT INTO ACCOUNTS (ID, BALANCE)
        VALUES (:WS-ID, :WS-AMT)
    END-EXEC.
    IF SQLCODE = 0
        EXEC SQL COMMIT END-EXEC
        DISPLAY "Committed"
    ELSE
        EXEC SQL ROLLBACK END-EXEC
        DISPLAY "Rolled back"
    END-IF.
"#));
}

// ═══════════════════════════════════════════════════════════
// COMPLEX SQL PROGRAMS
// ═══════════════════════════════════════════════════════════
#[test]
fn banking_transaction() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. BANKING.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DSN     PIC X(100) VALUE "sqlite:bank.db".
01 WS-FROM-ID PIC 9(10) VALUE 1001.
01 WS-TO-ID   PIC 9(10) VALUE 1002.
01 WS-AMOUNT  PIC 9(10)V99 VALUE 500.00.
01 WS-BAL     PIC 9(10)V99 VALUE 0.
01 SQLCODE    PIC S9(9) VALUE 0.
PROCEDURE DIVISION.
    EXEC SQL CONNECT :WS-DSN END-EXEC.
    EXEC SQL
        SELECT BALANCE INTO :WS-BAL
        FROM ACCOUNTS
        WHERE ACCOUNT_ID = :WS-FROM-ID
    END-EXEC.
    IF WS-BAL >= WS-AMOUNT
        EXEC SQL
            UPDATE ACCOUNTS
            SET BALANCE = BALANCE - :WS-AMOUNT
            WHERE ACCOUNT_ID = :WS-FROM-ID
        END-EXEC
        EXEC SQL
            UPDATE ACCOUNTS
            SET BALANCE = BALANCE + :WS-AMOUNT
            WHERE ACCOUNT_ID = :WS-TO-ID
        END-EXEC
        EXEC SQL COMMIT END-EXEC
        DISPLAY "Transfer complete"
    ELSE
        DISPLAY "Insufficient funds"
    END-IF.
    STOP RUN.
"#);
}

#[test]
fn customer_report() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. CUSTREPORT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DSN    PIC X(100) VALUE "sqlite:customers.db".
01 WS-ID     PIC 9(10).
01 WS-NAME   PIC X(50).
01 WS-CITY   PIC X(30).
01 WS-BAL    PIC 9(10)V99.
01 WS-TOTAL  PIC 9(12)V99 VALUE 0.
01 WS-COUNT  PIC 9(5) VALUE 0.
01 SQLCODE   PIC S9(9) VALUE 0.
PROCEDURE DIVISION.
    EXEC SQL CONNECT :WS-DSN END-EXEC.
    EXEC SQL
        DECLARE REPORT-CURSOR CURSOR FOR
        SELECT ID, NAME, CITY, BALANCE
        FROM CUSTOMERS
        ORDER BY NAME
    END-EXEC.
    EXEC SQL OPEN REPORT-CURSOR END-EXEC.
    DISPLAY "Customer Report".
    DISPLAY "========================================".
    PERFORM UNTIL SQLCODE NOT = 0
        EXEC SQL
            FETCH REPORT-CURSOR
            INTO :WS-ID, :WS-NAME, :WS-CITY, :WS-BAL
        END-EXEC
        IF SQLCODE = 0
            DISPLAY WS-ID " " WS-NAME " " WS-CITY " " WS-BAL
            ADD WS-BAL TO WS-TOTAL
            ADD 1 TO WS-COUNT
        END-IF
    END-PERFORM.
    EXEC SQL CLOSE REPORT-CURSOR END-EXEC.
    DISPLAY "========================================".
    DISPLAY "Total Customers: " WS-COUNT.
    DISPLAY "Total Balance:   " WS-TOTAL.
    STOP RUN.
"#);
}

#[test]
fn batch_insert() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. BATCHINS.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DSN   PIC X(100) VALUE "sqlite:test.db".
01 WS-I     PIC 9(5) VALUE 0.
01 WS-NAME  PIC X(20).
01 WS-VALUE PIC 9(8)V99 VALUE 0.
01 SQLCODE  PIC S9(9) VALUE 0.
PROCEDURE DIVISION.
    EXEC SQL CONNECT :WS-DSN END-EXEC.
    EXEC SQL
        DELETE FROM TEST_TABLE WHERE 1=1
    END-EXEC.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 100
        MOVE "Record" TO WS-NAME
        COMPUTE WS-VALUE = WS-I * 10.50
        EXEC SQL
            INSERT INTO TEST_TABLE (ID, NAME, VALUE)
            VALUES (:WS-I, :WS-NAME, :WS-VALUE)
        END-EXEC
    END-PERFORM.
    EXEC SQL COMMIT END-EXEC.
    DISPLAY "Inserted 100 records".
    STOP RUN.
"#);
}

#[test]
fn error_handling_sql() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. SQLERR.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DSN    PIC X(100) VALUE "sqlite:test.db".
01 WS-NAME   PIC X(50).
01 WS-ID     PIC 9(10) VALUE 99999.
01 SQLCODE   PIC S9(9) VALUE 0.
PROCEDURE DIVISION.
    EXEC SQL CONNECT :WS-DSN END-EXEC.
    EXEC SQL
        SELECT NAME INTO :WS-NAME
        FROM USERS WHERE ID = :WS-ID
    END-EXEC.
    EVALUATE SQLCODE
        WHEN 0
            DISPLAY "Found: " WS-NAME
        WHEN 100
            DISPLAY "No data found"
        WHEN OTHER
            DISPLAY "SQL Error: " SQLCODE
    END-EVALUATE.
    STOP RUN.
"#);
}

#[test]
fn multiple_tables() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. MULTITBL.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DSN      PIC X(100) VALUE "sqlite:shop.db".
01 WS-CUST-ID  PIC 9(10).
01 WS-CUST-NAME PIC X(50).
01 WS-ORD-ID   PIC 9(10).
01 WS-ORD-AMT  PIC 9(10)V99.
01 SQLCODE     PIC S9(9) VALUE 0.
PROCEDURE DIVISION.
    EXEC SQL CONNECT :WS-DSN END-EXEC.
    EXEC SQL
        SELECT C.NAME, O.ORDER_ID, O.AMOUNT
        INTO :WS-CUST-NAME, :WS-ORD-ID, :WS-ORD-AMT
        FROM CUSTOMERS C
        JOIN ORDERS O ON C.ID = O.CUSTOMER_ID
        WHERE C.ID = :WS-CUST-ID
    END-EXEC.
    IF SQLCODE = 0
        DISPLAY "Customer: " WS-CUST-NAME
        DISPLAY "Order: " WS-ORD-ID " Amount: " WS-ORD-AMT
    END-IF.
    STOP RUN.
"#);
}
