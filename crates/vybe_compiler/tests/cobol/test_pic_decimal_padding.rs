use super::helpers::{compile_ok, compile_ok_check, parse_ok};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

// ═══════════════════════════════════════════════════════════
// PIC EDITING — Formatted Display
// ═══════════════════════════════════════════════════════════

// ── Z (zero suppression) ───────────────────────────────────
#[test]
fn pic_z_basic() {
    compile_ok(&p("01 WS-X PIC Z(5)9 VALUE 42.", "    DISPLAY WS-X."));
}
#[test]
fn pic_zz99() {
    compile_ok(&p("01 WS-X PIC ZZ99 VALUE 42.", "    DISPLAY WS-X."));
}
#[test]
fn pic_zzz_all_zero() {
    compile_ok(&p("01 WS-X PIC Z(5) VALUE 0.", "    DISPLAY WS-X."));
}

// ── $ (currency) ───────────────────────────────────────────
#[test]
fn pic_dollar() {
    compile_ok(&p(
        "01 WS-AMT PIC $9(5).99 VALUE 1234.56.",
        "    DISPLAY WS-AMT.",
    ));
}
#[test]
fn pic_dollar_suppress() {
    compile_ok(&p(
        "01 WS-AMT PIC $ZZ,ZZ9.99 VALUE 1234.56.",
        "    DISPLAY WS-AMT.",
    ));
}
#[test]
fn pic_dollar_large() {
    compile_ok(&p(
        "01 WS-AMT PIC $$$,$$9.99 VALUE 75000.00.",
        "    DISPLAY WS-AMT.",
    ));
}

// ── , (comma insertion) ────────────────────────────────────
#[test]
fn pic_comma() {
    compile_ok(&p(
        "01 WS-X PIC 9(3),9(3) VALUE 123456.",
        "    DISPLAY WS-X.",
    ));
}
#[test]
fn pic_comma_z() {
    compile_ok(&p("01 WS-X PIC ZZZ,ZZ9 VALUE 12345.", "    DISPLAY WS-X."));
}

// ── . (decimal point) ──────────────────────────────────────
#[test]
fn pic_decimal() {
    compile_ok(&p("01 WS-X PIC 9(5).99 VALUE 123.45.", "    DISPLAY WS-X."));
}
#[test]
fn pic_v_decimal() {
    compile_ok(&p("01 WS-X PIC 9(5)V99 VALUE 123.45.", "    DISPLAY WS-X."));
}

// ── - / + (sign display) ──────────────────────────────────
#[test]
fn pic_minus_trail() {
    compile_ok(&p("01 WS-X PIC 9(5)- VALUE -100.", "    DISPLAY WS-X."));
}
#[test]
fn pic_plus_lead() {
    compile_ok(&p("01 WS-X PIC +9(5) VALUE 100.", "    DISPLAY WS-X."));
}
#[test]
fn pic_s_signed() {
    compile_ok(&p("01 WS-X PIC S9(5) VALUE -500.", "    DISPLAY WS-X."));
}

// ── * (asterisk fill / check protection) ──────────────────
#[test]
fn pic_asterisk() {
    compile_ok(&p(
        "01 WS-X PIC **(5)9.99 VALUE 42.50.",
        "    DISPLAY WS-X.",
    ));
}

// ── B (blank insertion) ────────────────────────────────────
#[test]
fn pic_blank() {
    compile_ok(&p(
        "01 WS-X PIC 9(3)B9(3) VALUE 123456.",
        "    DISPLAY WS-X.",
    ));
}

// ── / (slash insertion) ────────────────────────────────────
#[test]
fn pic_slash_date() {
    compile_ok(&p(
        "01 WS-DATE PIC 99/99/9999 VALUE 12252024.",
        "    DISPLAY WS-DATE.",
    ));
}

// ── 0 (zero insertion) ─────────────────────────────────────
#[test]
fn pic_zero_ins() {
    compile_ok(&p(
        "01 WS-X PIC 9(3)09(3) VALUE 123456.",
        "    DISPLAY WS-X.",
    ));
}

// ═══════════════════════════════════════════════════════════
// COMP-3 PACKED DECIMAL — Exact Arithmetic
// ═══════════════════════════════════════════════════════════

#[test]
fn comp3_add() {
    compile_ok(&p(
        "01 WS-A PIC 9(5)V99 COMP-3 VALUE 100.50.\n01 WS-B PIC 9(5)V99 COMP-3 VALUE 200.75.\n01 WS-C PIC 9(5)V99 COMP-3 VALUE 0.",
        "    ADD WS-A TO WS-B.\n    DISPLAY WS-B.",
    ));
}

#[test]
fn comp3_multiply() {
    compile_ok(&p(
        "01 WS-PRICE PIC 9(5)V99 COMP-3 VALUE 19.99.\n01 WS-QTY PIC 9(3) COMP-3 VALUE 5.\n01 WS-TOTAL PIC 9(7)V99 COMP-3 VALUE 0.",
        "    MULTIPLY WS-PRICE BY WS-QTY GIVING WS-TOTAL.\n    DISPLAY WS-TOTAL.",
    ));
}

#[test]
fn comp3_divide() {
    compile_ok(&p(
        "01 WS-AMT PIC 9(7)V99 COMP-3 VALUE 100.00.\n01 WS-PARTS PIC 9(3) COMP-3 VALUE 3.\n01 WS-EACH PIC 9(7)V99 COMP-3 VALUE 0.",
        "    DIVIDE WS-AMT BY WS-PARTS GIVING WS-EACH.\n    DISPLAY WS-EACH.",
    ));
}

#[test]
fn comp3_compute() {
    compile_ok(&p(
        "01 WS-PRINCIPAL PIC 9(8)V99 COMP-3 VALUE 10000.00.\n01 WS-RATE PIC 9(2)V99 COMP-3 VALUE 5.50.\n01 WS-INTEREST PIC 9(8)V99 COMP-3 VALUE 0.",
        "    COMPUTE WS-INTEREST = WS-PRINCIPAL * (WS-RATE / 100).\n    DISPLAY WS-INTEREST.",
    ));
}

#[test]
fn comp3_rounding() {
    compile_ok(&p(
        "01 WS-A PIC 9(5)V99 COMP-3 VALUE 0.\n01 WS-B PIC 9(5)V99 COMP-3 VALUE 10.00.\n01 WS-C PIC 9(5)V99 COMP-3 VALUE 3.",
        "    COMPUTE WS-A = WS-B / WS-C.\n    DISPLAY WS-A.",
    ));
}

// ── USAGE COMP / COMP-5 / BINARY ──────────────────────────
#[test]
fn comp_binary() {
    compile_ok(&p(
        "01 WS-X PIC 9(9) COMP VALUE 12345.",
        "    DISPLAY WS-X.",
    ));
}
#[test]
fn comp5() {
    compile_ok(&p(
        "01 WS-X PIC 9(9) USAGE BINARY VALUE 255.",
        "    DISPLAY WS-X.",
    ));
}

// ═══════════════════════════════════════════════════════════
// MOVE SPACE-PADDING SEMANTICS
// ═══════════════════════════════════════════════════════════

// ── Alpha: right-fill with spaces ──────────────────────────
#[test]
fn move_alpha_pad() {
    compile_ok(&p(
        "01 WS-A PIC X(20) VALUE \"Hello\".",
        "    DISPLAY WS-A.",
    ));
}

#[test]
fn move_alpha_short() {
    compile_ok(&p(
        "01 WS-SRC PIC X(5) VALUE \"Hi\".\n01 WS-DST PIC X(10).",
        "    MOVE WS-SRC TO WS-DST.\n    DISPLAY WS-DST.",
    ));
}

#[test]
fn move_alpha_truncate() {
    compile_ok(&p(
        "01 WS-SRC PIC X(20) VALUE \"Hello World\".\n01 WS-DST PIC X(5).",
        "    MOVE WS-SRC TO WS-DST.\n    DISPLAY WS-DST.",
    ));
}

// ── Numeric: left-fill with zeros ──────────────────────────
#[test]
fn move_num_zero_fill() {
    compile_ok(&p(
        "01 WS-SRC PIC 9(3) VALUE 42.\n01 WS-DST PIC 9(5).",
        "    MOVE WS-SRC TO WS-DST.\n    DISPLAY WS-DST.",
    ));
}

#[test]
fn move_num_to_larger() {
    compile_ok(&p(
        "01 WS-SRC PIC 9(3) VALUE 5.\n01 WS-DST PIC 9(8).",
        "    MOVE WS-SRC TO WS-DST.\n    DISPLAY WS-DST.",
    ));
}

// ── Numeric to alpha ───────────────────────────────────────
#[test]
fn move_num_to_alpha() {
    compile_ok(&p(
        "01 WS-NUM PIC 9(5) VALUE 12345.\n01 WS-STR PIC X(10).",
        "    MOVE WS-NUM TO WS-STR.\n    DISPLAY WS-STR.",
    ));
}

// ── Spaces/Zeros ───────────────────────────────────────────
#[test]
fn move_spaces_to_alpha() {
    compile_ok(&p(
        "01 WS-X PIC X(20) VALUE \"Old data\".",
        "    MOVE SPACES TO WS-X.\n    DISPLAY WS-X.",
    ));
}

#[test]
fn move_zeros_to_num() {
    compile_ok(&p(
        "01 WS-X PIC 9(5) VALUE 999.",
        "    MOVE ZEROS TO WS-X.\n    DISPLAY WS-X.",
    ));
}

// ═══════════════════════════════════════════════════════════
// COMPLEX PROGRAMS WITH FORMATTING
// ═══════════════════════════════════════════════════════════

#[test]
fn financial_report() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. FINRPT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PRICE    PIC 9(5)V99 COMP-3 VALUE 0.
01 WS-QTY      PIC 9(3)    VALUE 0.
01 WS-SUBTOTAL PIC 9(7)V99 COMP-3 VALUE 0.
01 WS-TAX      PIC 9(7)V99 COMP-3 VALUE 0.
01 WS-TOTAL    PIC 9(8)V99 COMP-3 VALUE 0.
01 WS-I        PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    MOVE 0 TO WS-TOTAL.
    MOVE 29.99 TO WS-PRICE.
    MOVE 3 TO WS-QTY.
    COMPUTE WS-SUBTOTAL = WS-PRICE * WS-QTY.
    COMPUTE WS-TAX = WS-SUBTOTAL * 0.08.
    COMPUTE WS-TOTAL = WS-SUBTOTAL + WS-TAX.
    DISPLAY "Subtotal: " WS-SUBTOTAL.
    DISPLAY "Tax:      " WS-TAX.
    DISPLAY "Total:    " WS-TOTAL.
    STOP RUN.
"#,
    );
}

#[test]
fn padded_record_output() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. PADREC.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RECORD.
   05 WS-ID     PIC 9(5)  VALUE 0.
   05 WS-NAME   PIC X(30) VALUE SPACES.
   05 WS-AMOUNT PIC 9(8)V99 VALUE 0.
PROCEDURE DIVISION.
    MOVE 12345 TO WS-ID.
    MOVE "John Smith" TO WS-NAME.
    MOVE 5000.50 TO WS-AMOUNT.
    DISPLAY WS-ID " " WS-NAME " " WS-AMOUNT.
    STOP RUN.
"#,
    );
}

#[test]
fn currency_formatting() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. CURR.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PRICE PIC 9(5)V99 VALUE 1234.56.
01 WS-TOTAL PIC 9(8)V99 VALUE 0.
01 WS-QTY   PIC 9(3) VALUE 10.
PROCEDURE DIVISION.
    COMPUTE WS-TOTAL = WS-PRICE * WS-QTY.
    DISPLAY "Unit Price: " WS-PRICE.
    DISPLAY "Quantity:   " WS-QTY.
    DISPLAY "Total:      " WS-TOTAL.
    STOP RUN.
"#,
    );
}

#[test]
fn mixed_comp_types() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. MIXCOMP.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(5) COMP VALUE 100.
01 WS-B PIC 9(5) COMP-3 VALUE 200.
01 WS-C PIC 9(5) USAGE BINARY VALUE 300.
01 WS-D PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    COMPUTE WS-D = WS-A + WS-B + WS-C.
    DISPLAY "Result: " WS-D.
    STOP RUN.
"#,
    );
}
