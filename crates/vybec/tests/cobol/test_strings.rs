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

// ── STRING (concatenation) ─────────────────────────────────
#[test] fn string_two() { compile_ok(&p(
    "01 A PIC X(10) VALUE \"Hello\".\n01 B PIC X(10) VALUE \"World\".\n01 R PIC X(25).",
    "    STRING A DELIMITED BY SPACE \" \" DELIMITED BY SIZE B DELIMITED BY SPACE INTO R."
)); }
#[test] fn string_three() { compile_ok(&p(
    "01 A PIC X(5) VALUE \"A\".\n01 B PIC X(5) VALUE \"B\".\n01 C PIC X(5) VALUE \"C\".\n01 R PIC X(20).",
    "    STRING A DELIMITED BY SIZE B DELIMITED BY SIZE C DELIMITED BY SIZE INTO R."
)); }
#[test] fn string_literal() { compile_ok(&p(
    "01 R PIC X(30).",
    "    STRING \"Hello\" DELIMITED BY SIZE \" World\" DELIMITED BY SIZE INTO R."
)); }

// ── UNSTRING (split) ───────────────────────────────────────
#[test] fn unstring_comma() { compile_ok(&p(
    "01 SRC PIC X(30) VALUE \"A,B,C\".\n01 F1 PIC X(10).\n01 F2 PIC X(10).\n01 F3 PIC X(10).",
    "    UNSTRING SRC DELIMITED BY \",\" INTO F1 F2 F3."
)); }
#[test] fn unstring_space() { compile_ok(&p(
    "01 SRC PIC X(30) VALUE \"Hello World Cobol\".\n01 W1 PIC X(10).\n01 W2 PIC X(10).\n01 W3 PIC X(10).",
    "    UNSTRING SRC DELIMITED BY \" \" INTO W1 W2 W3."
)); }

// ── INSPECT TALLYING ───────────────────────────────────────
#[test] fn inspect_tally_all() { compile_ok(&p(
    "01 TXT PIC X(20) VALUE \"Hello World\".\n01 CNT PIC 9(3) VALUE 0.",
    "    INSPECT TXT TALLYING CNT FOR ALL \"l\"."
)); }
#[test] fn inspect_tally_leading() { compile_ok(&p(
    "01 TXT PIC X(20) VALUE \"000123\".\n01 CNT PIC 9(3) VALUE 0.",
    "    INSPECT TXT TALLYING CNT FOR LEADING \"0\"."
)); }

// ── INSPECT REPLACING ──────────────────────────────────────
#[test] fn inspect_replace_all() { compile_ok(&p(
    "01 TXT PIC X(20) VALUE \"Hello World\".",
    "    INSPECT TXT REPLACING ALL \"l\" BY \"r\"."
)); }
#[test] fn inspect_replace_first() { compile_ok(&p(
    "01 TXT PIC X(20) VALUE \"aabaa\".",
    "    INSPECT TXT REPLACING FIRST \"a\" BY \"X\"."
)); }

// ── Reference Modification ─────────────────────────────────
#[test] fn refmod_basic() { compile_ok(&p(
    "01 TXT PIC X(20) VALUE \"Hello World\".\n01 SUB PIC X(5).",
    "    MOVE TXT(1:5) TO SUB.\n    DISPLAY SUB."
)); }
#[test] fn refmod_middle() { compile_ok(&p(
    "01 TXT PIC X(20) VALUE \"Hello World\".\n01 SUB PIC X(5).",
    "    MOVE TXT(7:5) TO SUB.\n    DISPLAY SUB."
)); }

// ── FUNCTION UPPER-CASE / LOWER-CASE ──────────────────────
#[test] fn func_upper() { compile_ok(&p(
    "01 TXT PIC X(20) VALUE \"hello\".\n01 R PIC X(20).",
    "    MOVE FUNCTION UPPER-CASE(TXT) TO R.\n    DISPLAY R."
)); }
#[test] fn func_lower() { compile_ok(&p(
    "01 TXT PIC X(20) VALUE \"HELLO\".\n01 R PIC X(20).",
    "    MOVE FUNCTION LOWER-CASE(TXT) TO R.\n    DISPLAY R."
)); }
#[test] fn func_trim() { compile_ok(&p(
    "01 TXT PIC X(20) VALUE \"  hello  \".\n01 R PIC X(20).",
    "    MOVE FUNCTION TRIM(TXT) TO R.\n    DISPLAY R."
)); }
#[test] fn func_reverse() { compile_ok(&p(
    "01 TXT PIC X(10) VALUE \"Hello\".\n01 R PIC X(10).",
    "    MOVE FUNCTION REVERSE(TXT) TO R.\n    DISPLAY R."
)); }
#[test] fn func_length() { compile_ok(&p(
    "01 TXT PIC X(20) VALUE \"Hello\".\n01 L PIC 9(5).",
    "    MOVE FUNCTION LENGTH(TXT) TO L.\n    DISPLAY L."
)); }
#[test] fn func_substitute() { compile_ok(&p(
    "01 TXT PIC X(30) VALUE \"Hello World\".\n01 R PIC X(30).",
    "    MOVE FUNCTION SUBSTITUTE(TXT \"World\" \"COBOL\") TO R.\n    DISPLAY R."
)); }
#[test] fn func_concatenate() { compile_ok(&p(
    "01 A PIC X(10) VALUE \"Hello\".\n01 B PIC X(10) VALUE \"World\".\n01 R PIC X(25).",
    "    MOVE FUNCTION CONCATENATE(A B) TO R.\n    DISPLAY R."
)); }
