use vybec::parser_cobol::parse;
use vybec::compiler_cobol::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

fn parse_ok(src: &str) -> bool { parse(src).is_ok() }

// ── PIC Clause Varieties ───────────────────────────────────
#[test] fn pic_x() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-A PIC X(10).\nPROCEDURE DIVISION.\n    STOP RUN."); }
#[test] fn pic_x_value() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-A PIC X(10) VALUE \"Hello\".\nPROCEDURE DIVISION.\n    STOP RUN."); }
#[test] fn pic_9() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-A PIC 9(5).\nPROCEDURE DIVISION.\n    STOP RUN."); }
#[test] fn pic_9_value() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-A PIC 9(5) VALUE 12345.\nPROCEDURE DIVISION.\n    STOP RUN."); }
#[test] fn pic_s9() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-A PIC S9(5) VALUE -100.\nPROCEDURE DIVISION.\n    STOP RUN."); }
#[test] fn pic_9v99() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-A PIC 9(5)V99 VALUE 123.45.\nPROCEDURE DIVISION.\n    STOP RUN."); }
#[test] fn pic_a() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-A PIC A(10) VALUE \"Hello\".\nPROCEDURE DIVISION.\n    STOP RUN."); }
#[test] fn value_spaces() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-A PIC X(10) VALUE SPACES.\nPROCEDURE DIVISION.\n    STOP RUN."); }
#[test] fn value_zeros() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-A PIC 9(5) VALUE ZEROS.\nPROCEDURE DIVISION.\n    STOP RUN."); }
#[test] fn value_low_values() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-A PIC X(5) VALUE LOW-VALUES.\nPROCEDURE DIVISION.\n    STOP RUN."); }
#[test] fn value_high_values() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-A PIC X(5) VALUE HIGH-VALUES.\nPROCEDURE DIVISION.\n    STOP RUN."); }

// ── Group Items ────────────────────────────────────────────
#[test] fn group_2_levels() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-REC.\n   05 WS-A PIC X(10).\n   05 WS-B PIC 9(5).\nPROCEDURE DIVISION.\n    STOP RUN."); }
#[test] fn group_3_levels() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-REC.\n   05 WS-GRP.\n      10 WS-A PIC X(5).\n      10 WS-B PIC X(5).\n   05 WS-C PIC 9(3).\nPROCEDURE DIVISION.\n    STOP RUN."); }
#[test] fn group_with_values() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-REC.\n   05 WS-NAME PIC X(20) VALUE \"John\".\n   05 WS-AGE PIC 9(3) VALUE 25.\n   05 WS-CITY PIC X(20) VALUE \"NYC\".\nPROCEDURE DIVISION.\n    STOP RUN."); }
#[test] fn multiple_01_items() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-A PIC X(10).\n01 WS-B PIC 9(5).\n01 WS-C PIC X(20).\nPROCEDURE DIVISION.\n    STOP RUN."); }

// ── 88-Level Conditions ────────────────────────────────────
#[test] fn cond_88_single() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-STATUS PIC X(1).\n   88 ACTIVE VALUE \"A\".\nPROCEDURE DIVISION.\n    STOP RUN."); }
#[test] fn cond_88_multiple_values() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-STATUS PIC X(1).\n   88 ACTIVE VALUE \"A\".\n   88 INACTIVE VALUE \"I\".\n   88 PENDING VALUE \"P\".\nPROCEDURE DIVISION.\n    STOP RUN."); }
#[test] fn cond_88_numeric() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-CODE PIC 9(1).\n   88 IS-YES VALUE 1.\n   88 IS-NO VALUE 0.\nPROCEDURE DIVISION.\n    STOP RUN."); }

// ── OCCURS ─────────────────────────────────────────────────
#[test] fn occurs_basic() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-TABLE.\n   05 WS-ITEM PIC X(10) OCCURS 10 TIMES.\nPROCEDURE DIVISION.\n    STOP RUN."); }
#[test] fn occurs_nested() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-TABLE.\n   05 WS-ROW OCCURS 5 TIMES.\n      10 WS-COL PIC 9(3) OCCURS 3 TIMES.\nPROCEDURE DIVISION.\n    STOP RUN."); }
#[test] fn occurs_group() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-EMPLOYEES.\n   05 WS-EMP OCCURS 100 TIMES.\n      10 WS-EMP-NAME PIC X(30).\n      10 WS-EMP-AGE PIC 9(3).\nPROCEDURE DIVISION.\n    STOP RUN."); }

// ── Level 77 (standalone) ──────────────────────────────────
#[test] fn level_77() { compile_ok("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n77 WS-COUNTER PIC 9(5) VALUE 0.\nPROCEDURE DIVISION.\n    STOP RUN."); }
