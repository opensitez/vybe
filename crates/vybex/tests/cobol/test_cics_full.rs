use super::helpers::{compile_ok, parse_ok, compile_ok_check};



fn p(data: &str, body: &str) -> String {
    format!("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.", data, body)
}

fn d() -> &'static str { "01 WS-DATA PIC X(100).\n01 WS-KEY PIC X(20).\n01 WS-USER PIC X(20).\n01 WS-TERM PIC X(8).\n01 WS-SYSID PIC X(4).\n01 WS-PTR PIC X(10).\n01 WS-MSG PIC X(80)." }

// ═══════════════════════════════════════════════════════════
// HANDLE CONDITION / HANDLE AID
// ═══════════════════════════════════════════════════════════
#[test] fn handle_condition() { compile_ok(&p(d(), "    EXEC CICS HANDLE CONDITION ERROR(ERROR-PARA) NOTFND(NOTFOUND-PARA) END-EXEC.")); }
#[test] fn handle_condition_multi() { compile_ok(&p(d(), "    EXEC CICS HANDLE CONDITION DUPREC(DUP-PARA) LENGERR(LEN-PARA) INVREQ(INV-PARA) END-EXEC.")); }
#[test] fn handle_aid() { compile_ok(&p(d(), "    EXEC CICS HANDLE AID PF1(HELP-PARA) PF3(EXIT-PARA) ENTER(PROCESS-PARA) CLEAR(CLEAR-PARA) END-EXEC.")); }

// ═══════════════════════════════════════════════════════════
// WRITEQ TS / READQ TS / DELETEQ TS (Temporary Storage)
// ═══════════════════════════════════════════════════════════
#[test] fn writeq_ts() { compile_ok(&p(d(), "    EXEC CICS WRITEQ TS QUEUE(MYQUEUE) FROM(WS-DATA) END-EXEC.")); }
#[test] fn readq_ts() { compile_ok(&p(d(), "    EXEC CICS READQ TS QUEUE(MYQUEUE) INTO(WS-DATA) END-EXEC.")); }
#[test] fn readq_ts_item() { compile_ok(&p(d(), "    EXEC CICS READQ TS QUEUE(MYQUEUE) INTO(WS-DATA) ITEM(3) END-EXEC.")); }
#[test] fn deleteq_ts() { compile_ok(&p(d(), "    EXEC CICS DELETEQ TS QUEUE(MYQUEUE) END-EXEC.")); }
#[test] fn writeq_td() { compile_ok(&p(d(), "    EXEC CICS WRITEQ TD QUEUE(ERRLOG) FROM(WS-MSG) END-EXEC.")); }
#[test] fn readq_td() { compile_ok(&p(d(), "    EXEC CICS READQ TD QUEUE(ERRLOG) INTO(WS-MSG) END-EXEC.")); }

// ═══════════════════════════════════════════════════════════
// ENQ / DEQ (Resource Locking)
// ═══════════════════════════════════════════════════════════
#[test] fn enq_resource() { compile_ok(&p(d(), "    EXEC CICS ENQ RESOURCE(CUSTLOCK) LENGTH(8) END-EXEC.")); }
#[test] fn deq_resource() { compile_ok(&p(d(), "    EXEC CICS DEQ RESOURCE(CUSTLOCK) LENGTH(8) END-EXEC.")); }
#[test] fn enq_deq_pair() { compile_ok(&p(d(),
    "    EXEC CICS ENQ RESOURCE(ACCTLOCK) END-EXEC.\n    MOVE \"Updated\" TO WS-DATA.\n    EXEC CICS DEQ RESOURCE(ACCTLOCK) END-EXEC."
)); }

// ═══════════════════════════════════════════════════════════
// ASSIGN (System Information)
// ═══════════════════════════════════════════════════════════
#[test] fn assign_userid() { compile_ok(&p(d(), "    EXEC CICS ASSIGN USERID(WS-USER) END-EXEC.")); }
#[test] fn assign_sysid() { compile_ok(&p(d(), "    EXEC CICS ASSIGN SYSID(WS-SYSID) END-EXEC.")); }
#[test] fn assign_multi() { compile_ok(&p(d(), "    EXEC CICS ASSIGN USERID(WS-USER) SYSID(WS-SYSID) END-EXEC.")); }

// ═══════════════════════════════════════════════════════════
// PUT / GET CONTAINER (Channels)
// ═══════════════════════════════════════════════════════════
#[test] fn put_container() { compile_ok(&p(d(), "    EXEC CICS PUT CONTAINER(CUSTDATA) CHANNEL(MYCHANNEL) FROM(WS-DATA) END-EXEC.")); }
#[test] fn get_container() { compile_ok(&p(d(), "    EXEC CICS GET CONTAINER(CUSTDATA) CHANNEL(MYCHANNEL) INTO(WS-DATA) END-EXEC.")); }
#[test] fn put_get_roundtrip() { compile_ok(&p(d(),
    "    EXEC CICS PUT CONTAINER(ITEM1) CHANNEL(CH1) FROM(WS-DATA) END-EXEC.\n    EXEC CICS GET CONTAINER(ITEM1) CHANNEL(CH1) INTO(WS-MSG) END-EXEC."
)); }

// ═══════════════════════════════════════════════════════════
// DELAY / START / SUSPEND / POST / WAIT
// ═══════════════════════════════════════════════════════════
#[test] fn delay_seconds() { compile_ok(&p(d(), "    EXEC CICS DELAY SECONDS(5) END-EXEC.")); }
#[test] fn delay_interval() { compile_ok(&p(d(), "    EXEC CICS DELAY INTERVAL(001000) END-EXEC.")); }
#[test] fn start_trans() { compile_ok(&p(d(), "    EXEC CICS START TRANSID(NXTTX) END-EXEC.")); }
#[test] fn suspend_task() { compile_ok(&p(d(), "    EXEC CICS SUSPEND END-EXEC.")); }
#[test] fn post_event() { compile_ok(&p(d(), "    EXEC CICS POST EVENT(MYEVENT) END-EXEC.")); }
#[test] fn wait_event() { compile_ok(&p(d(), "    EXEC CICS WAIT EVENT END-EXEC.")); }

// ═══════════════════════════════════════════════════════════
// WEB (CICS Web Services)
// ═══════════════════════════════════════════════════════════
#[test] fn web_send() { compile_ok(&p(d(), "    EXEC CICS WEB SEND FROM(WS-DATA) LENGTH(100) END-EXEC.")); }
#[test] fn web_receive() { compile_ok(&p(d(), "    EXEC CICS WEB RECEIVE INTO(WS-DATA) LENGTH(100) END-EXEC.")); }

// ═══════════════════════════════════════════════════════════
// DOCUMENT
// ═══════════════════════════════════════════════════════════
#[test] fn document_create() { compile_ok(&p(d(), "    EXEC CICS DOCUMENT CREATE DOCTOKEN(WS-PTR) END-EXEC.")); }

// ═══════════════════════════════════════════════════════════
// ABEND
// ═══════════════════════════════════════════════════════════
#[test] fn abend_code() { compile_ok(&p(d(), "    EXEC CICS ABEND ABCODE(ASRA) END-EXEC.")); }
#[test] fn abend_custom() { compile_ok(&p(d(), "    EXEC CICS ABEND ABCODE(USR1) END-EXEC.")); }

// ═══════════════════════════════════════════════════════════
// INQUIRE / SET
// ═══════════════════════════════════════════════════════════
#[test] fn inquire_program() { compile_ok(&p(d(), "    EXEC CICS INQUIRE PROGRAM(MYPROG) END-EXEC.")); }

// ═══════════════════════════════════════════════════════════
// COMPLEX CICS PROGRAMS
// ═══════════════════════════════════════════════════════════
#[test]
fn cics_customer_inquiry() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. CUSTINQ.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-CUSTID  PIC 9(10).
01 WS-NAME    PIC X(30).
01 WS-BALANCE PIC 9(10)V99.
01 WS-INPUT   PIC X(80).
01 WS-OUTPUT  PIC X(200).
01 WS-USER    PIC X(8).
PROCEDURE DIVISION.
    EXEC CICS ASSIGN USERID(WS-USER) END-EXEC.
    EXEC CICS HANDLE CONDITION NOTFND(NOT-FOUND-PARA) END-EXEC.
    EXEC CICS RECEIVE MAP(INQMAP) INTO(WS-INPUT) END-EXEC.
    EXEC CICS READ FILE(CUSTFILE) INTO(WS-NAME) RIDFLD(WS-CUSTID) END-EXEC.
    EXEC CICS SEND MAP(DETMAP) FROM(WS-OUTPUT) END-EXEC.
    EXEC CICS RETURN TRANSID(CINQ) END-EXEC.
    STOP RUN.
NOT-FOUND-PARA.
    DISPLAY "Customer not found".
"#);
}

#[test]
fn cics_queue_processing() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. QPROC.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ITEM PIC X(100).
01 WS-COUNT PIC 9(5) VALUE 0.
01 WS-I PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 10
        MOVE "Item" TO WS-ITEM
        EXEC CICS WRITEQ TS QUEUE(WORKQ) FROM(WS-ITEM) END-EXEC
        ADD 1 TO WS-COUNT
    END-PERFORM.
    DISPLAY "Queued " WS-COUNT " items".
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 10
        EXEC CICS READQ TS QUEUE(WORKQ) INTO(WS-ITEM) END-EXEC
        DISPLAY "Read: " WS-ITEM
    END-PERFORM.
    EXEC CICS DELETEQ TS QUEUE(WORKQ) END-EXEC.
    STOP RUN.
"#);
}

#[test]
fn cics_channel_program() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. CHPROG.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REQUEST  PIC X(200).
01 WS-RESPONSE PIC X(500).
01 WS-USER     PIC X(8).
PROCEDURE DIVISION.
    EXEC CICS ASSIGN USERID(WS-USER) END-EXEC.
    MOVE "Get customer 1001" TO WS-REQUEST.
    EXEC CICS PUT CONTAINER(REQUEST) CHANNEL(SVCCHAN) FROM(WS-REQUEST) END-EXEC.
    EXEC CICS LINK PROGRAM(CUSTSERV) END-EXEC.
    EXEC CICS GET CONTAINER(RESPONSE) CHANNEL(SVCCHAN) INTO(WS-RESPONSE) END-EXEC.
    DISPLAY "Response: " WS-RESPONSE.
    EXEC CICS RETURN END-EXEC.
    STOP RUN.
"#);
}

#[test]
fn cics_with_locking() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. LOCKPROG.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-BALANCE PIC 9(10)V99 VALUE 0.
01 WS-AMOUNT  PIC 9(10)V99 VALUE 500.
01 WS-ACCTID  PIC X(10) VALUE "ACCT1001".
PROCEDURE DIVISION.
    EXEC CICS ENQ RESOURCE(WS-ACCTID) LENGTH(10) END-EXEC.
    EXEC CICS READ FILE(ACCTFILE) INTO(WS-BALANCE) RIDFLD(WS-ACCTID) END-EXEC.
    ADD WS-AMOUNT TO WS-BALANCE.
    EXEC CICS REWRITE FILE(ACCTFILE) FROM(WS-BALANCE) END-EXEC.
    EXEC CICS DEQ RESOURCE(WS-ACCTID) LENGTH(10) END-EXEC.
    DISPLAY "Updated balance: " WS-BALANCE.
    EXEC CICS RETURN END-EXEC.
    STOP RUN.
"#);
}

#[test]
fn cics_web_service() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. WEBSVC.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REQUEST  PIC X(500).
01 WS-RESPONSE PIC X(1000).
01 WS-DOCTOKEN PIC X(20).
PROCEDURE DIVISION.
    EXEC CICS WEB RECEIVE INTO(WS-REQUEST) LENGTH(500) END-EXEC.
    DISPLAY "Request: " WS-REQUEST.
    MOVE "Response data" TO WS-RESPONSE.
    EXEC CICS DOCUMENT CREATE DOCTOKEN(WS-DOCTOKEN) END-EXEC.
    EXEC CICS WEB SEND FROM(WS-RESPONSE) LENGTH(1000) END-EXEC.
    EXEC CICS RETURN END-EXEC.
    STOP RUN.
"#);
}

#[test]
fn cics_error_handling() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. ERRHNDL.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATA PIC X(100).
01 WS-KEY  PIC X(20) VALUE "CUST001".
PROCEDURE DIVISION.
    EXEC CICS HANDLE CONDITION
        NOTFND(NOT-FOUND)
        DUPREC(DUPLICATE)
        ERROR(GENERAL-ERROR)
    END-EXEC.
    EXEC CICS READ FILE(CUSTFILE) INTO(WS-DATA) RIDFLD(WS-KEY) END-EXEC.
    DISPLAY "Found: " WS-DATA.
    EXEC CICS RETURN END-EXEC.
    STOP RUN.
NOT-FOUND.
    DISPLAY "Record not found".
DUPLICATE.
    DISPLAY "Duplicate record".
GENERAL-ERROR.
    DISPLAY "An error occurred".
    EXEC CICS ABEND ABCODE(UERR) END-EXEC.
"#);
}
