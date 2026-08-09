use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

// ── DISPLAY ────────────────────────────────────────────────
#[test]
fn display_literal() {
    compile_ok(&p("", "    DISPLAY \"Hello World\"."));
}
#[test]
fn display_var() {
    compile_ok(&p("01 X PIC X(10) VALUE \"Test\".", "    DISPLAY X."));
}
#[test]
fn display_multi() {
    compile_ok(&p(
        "01 N PIC X(5) VALUE \"Bob\".\n01 A PIC 9(3) VALUE 30.",
        "    DISPLAY \"Name: \" N \" Age: \" A.",
    ));
}
#[test]
fn display_number() {
    compile_ok(&p("01 X PIC 9(5) VALUE 42.", "    DISPLAY X."));
}

// ── ACCEPT ─────────────────────────────────────────────────
#[test]
fn accept_basic() {
    compile_ok(&p("01 X PIC X(20).", "    ACCEPT X."));
}
#[test]
fn accept_from_date() {
    compile_ok(&p("01 D PIC X(8).", "    ACCEPT D FROM DATE."));
}
#[test]
fn accept_from_time() {
    compile_ok(&p("01 T PIC X(8).", "    ACCEPT T FROM TIME."));
}
#[test]
fn accept_from_day() {
    compile_ok(&p("01 D PIC X(5).", "    ACCEPT D FROM DAY."));
}

// ── MOVE ───────────────────────────────────────────────────
#[test]
fn move_literal_to_var() {
    compile_ok(&p("01 X PIC X(10).", "    MOVE \"Hello\" TO X."));
}
#[test]
fn move_num_to_var() {
    compile_ok(&p("01 X PIC 9(5).", "    MOVE 42 TO X."));
}
#[test]
fn move_var_to_var() {
    compile_ok(&p(
        "01 A PIC X(10) VALUE \"Hi\".\n01 B PIC X(10).",
        "    MOVE A TO B.",
    ));
}
#[test]
fn move_spaces() {
    compile_ok(&p("01 X PIC X(10).", "    MOVE SPACES TO X."));
}
#[test]
fn move_zeros() {
    compile_ok(&p("01 X PIC 9(5).", "    MOVE ZEROS TO X."));
}
#[test]
fn move_corresponding() {
    compile_ok(&p(
        "01 SRC.\n   05 WS-NAME PIC X(10) VALUE \"Alice\".\n   05 WS-AGE PIC 9(3) VALUE 30.\n01 DST.\n   05 WS-NAME PIC X(10).\n   05 WS-AGE PIC 9(3).",
        "    MOVE CORRESPONDING SRC TO DST.",
    ));
}

// ── INITIALIZE ─────────────────────────────────────────────
#[test]
fn initialize_var() {
    compile_ok(&p("01 X PIC X(10) VALUE \"Old\".", "    INITIALIZE X."));
}
#[test]
fn initialize_group() {
    compile_ok(&p(
        "01 REC.\n   05 A PIC X(10) VALUE \"Old\".\n   05 B PIC 9(5) VALUE 99.",
        "    INITIALIZE REC.",
    ));
}

// ── SET ────────────────────────────────────────────────────
#[test]
fn set_true() {
    compile_ok(&p(
        "01 WS-FLAG PIC 9(1).\n   88 IS-ON VALUE 1.",
        "    SET IS-ON TO TRUE.",
    ));
}
#[test]
fn set_false() {
    compile_ok(&p(
        "01 WS-FLAG PIC 9(1).\n   88 IS-OFF VALUE 0.",
        "    SET IS-OFF TO FALSE.",
    ));
}

// ── CALL ───────────────────────────────────────────────────
#[test]
fn call_basic() {
    compile_ok(&p("01 X PIC 9(5).", "    CALL \"SUBPROG\" USING X."));
}
#[test]
fn call_multi_args() {
    compile_ok(&p(
        "01 A PIC X(10).\n01 B PIC 9(5).",
        "    CALL \"PROCESS\" USING A B.",
    ));
}
#[test]
fn call_no_args() {
    compile_ok(&p("", "    CALL \"INIT\"."));
}
#[test]
fn call_with_returning_local() {
    compile_ok(&p(
        "01 RET PIC 9(3).",
        "    CALL \"SUBPROG\" RETURNING RET.",
    ));
}

// ── RAISE ──────────────────────────────────────────────────
#[test]
fn raise_string() {
    compile_ok(&p("", "    RAISE EXCEPTION \"Error occurred\"."));
}

// ── JSON ───────────────────────────────────────────────────
#[test]
fn json_gen() {
    compile_ok(&p(
        "01 REC.\n   05 NAME PIC X(10) VALUE \"Alice\".\n   05 AGE PIC 9(3) VALUE 30.\n01 J PIC X(100).",
        "    JSON GENERATE J FROM REC.",
    ));
}
#[test]
fn json_par() {
    compile_ok(&p(
        "01 J PIC X(100) VALUE '{\"name\":\"Bob\"}'.\n01 REC.\n   05 NAME PIC X(10).",
        "    JSON PARSE J INTO REC.",
    ));
}

// ── File I/O ───────────────────────────────────────────────
#[test]
fn open_input() {
    compile_ok(&p("", "    OPEN INPUT WS-FILE."));
}
#[test]
fn open_output() {
    compile_ok(&p("", "    OPEN OUTPUT WS-FILE."));
}
#[test]
fn close_file() {
    compile_ok(&p("", "    CLOSE WS-FILE."));
}
#[test]
fn read_file() {
    compile_ok(&p("01 REC PIC X(80).", "    READ WS-FILE INTO REC."));
}
#[test]
fn write_file() {
    compile_ok(&p(
        "01 REC PIC X(80) VALUE \"Data\".",
        "    WRITE WS-REC FROM REC.",
    ));
}

#[test]
fn sort_and_merge_program() {
    compile_ok(&p(
        "01 WS-KEY PIC 9(5).",
        "    SORT WS-FILE ON ASCENDING KEY WS-KEY.\n    MERGE WS-FILE ON DESCENDING KEY WS-KEY.",
    ));
}

#[test]
fn close_and_reopen_file_sequence() {
    compile_ok(&p(
        "",
        "    OPEN OUTPUT WS-FILE.\n    CLOSE WS-FILE.\n    OPEN INPUT WS-FILE.\n    CLOSE WS-FILE.",
    ));
}

// ── SORT ───────────────────────────────────────────────────
#[test]
fn sort_ascending() {
    compile_ok(&p("", "    SORT WS-FILE ON ASCENDING KEY WS-KEY."));
}
#[test]
fn sort_descending() {
    compile_ok(&p("", "    SORT WS-FILE ON DESCENDING KEY WS-KEY."));
}

// ── SEARCH ─────────────────────────────────────────────────
#[test] // SEARCH with subscript parsing needs work
fn search_basic() {
    compile_ok(&p(
        "01 TBL.\n   05 ITEM PIC X(10) OCCURS 10 TIMES.",
        "    SEARCH ITEM\n        AT END\n            DISPLAY \"Not found\"\n        WHEN ITEM(1) = \"A\"\n            DISPLAY \"Found\"\n    END-SEARCH.",
    ));
}
