use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn db_connect_query_disconnect_pattern_compiles() {
    compile_ok(&p(
        "01 WS-DSN PIC X(100) VALUE \"sqlite:test.db\".",
        "    EXEC SQL CONNECT :WS-DSN END-EXEC.\n    EXEC SQL SELECT 1 END-EXEC.\n    EXEC SQL COMMIT END-EXEC.",
    ));
}

#[test]
fn sql_cursor_lifecycle_compiles() {
    compile_ok(&p(
        "01 WS-ID PIC 9(5).",
        "    EXEC SQL DECLARE C1 CURSOR FOR SELECT ID FROM USERS END-EXEC.\n    EXEC SQL OPEN C1 END-EXEC.\n    EXEC SQL FETCH C1 INTO :WS-ID END-EXEC.\n    EXEC SQL CLOSE C1 END-EXEC.",
    ));
}

#[test]
fn db_transaction_pattern_compiles() {
    compile_ok(&p(
        "01 WS-DSN PIC X(80) VALUE \"sqlite:app.db\".",
        "    EXEC SQL CONNECT :WS-DSN END-EXEC.\n    EXEC SQL INSERT INTO LOGS(ID) VALUES(1) END-EXEC.\n    EXEC SQL COMMIT END-EXEC.",
    ));
}

#[test]
fn file_open_read_close_pattern_compiles() {
    compile_ok(&p(
        "01 WS-REC PIC X(80).",
        "    OPEN INPUT WS-FILE.\n    READ WS-FILE INTO WS-REC.\n    CLOSE WS-FILE.",
    ));
}
