use vybe_parser_cobol::parse;
use vybe_compiler_cobol::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

fn p(body: &str) -> String {
    format!("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 WS-A PIC 9(10) VALUE 0.\n01 WS-B PIC 9(10) VALUE 0.\n01 WS-C PIC 9(10) VALUE 0.\n01 WS-D PIC 9(10) VALUE 0.\n01 WS-R PIC 9(10) VALUE 0.\nPROCEDURE DIVISION.\n{}\n    STOP RUN.", body)
}

// ── ADD ────────────────────────────────────────────────────
#[test] fn add_to() { compile_ok(&p("    ADD 5 TO WS-A.")); }
#[test] fn add_giving() { compile_ok(&p("    ADD WS-A WS-B GIVING WS-C.")); }
#[test] fn add_literal_to() { compile_ok(&p("    ADD 10 TO WS-A.")); }
#[test] fn add_multiple() { compile_ok(&p("    ADD 1 2 3 TO WS-A.")); }

// ── SUBTRACT ───────────────────────────────────────────────
#[test] fn subtract_from() { compile_ok(&p("    SUBTRACT 5 FROM WS-A.")); }
#[test] fn subtract_giving() { compile_ok(&p("    SUBTRACT WS-B FROM WS-A GIVING WS-C.")); }
#[test] fn subtract_literal() { compile_ok(&p("    SUBTRACT 100 FROM WS-A.")); }

// ── MULTIPLY ───────────────────────────────────────────────
#[test] fn multiply_by() { compile_ok(&p("    MULTIPLY 3 BY WS-A.")); }
#[test] fn multiply_giving() { compile_ok(&p("    MULTIPLY WS-A BY WS-B GIVING WS-C.")); }

// ── DIVIDE ─────────────────────────────────────────────────
#[test] fn divide_giving() { compile_ok(&p("    DIVIDE WS-A BY 3 GIVING WS-C.")); }
#[test] fn divide_remainder() { compile_ok(&p("    DIVIDE 17 BY 5 GIVING WS-C REMAINDER WS-R.")); }

// ── COMPUTE ────────────────────────────────────────────────
#[test] fn compute_add() { compile_ok(&p("    COMPUTE WS-C = WS-A + WS-B.")); }
#[test] fn compute_sub() { compile_ok(&p("    COMPUTE WS-C = WS-A - WS-B.")); }
#[test] fn compute_mul() { compile_ok(&p("    COMPUTE WS-C = WS-A * WS-B.")); }
#[test] fn compute_div() { compile_ok(&p("    COMPUTE WS-C = WS-A / WS-B.")); }
#[test] fn compute_pow() { compile_ok(&p("    COMPUTE WS-C = WS-A ** 2.")); }
#[test] fn compute_complex() { compile_ok(&p("    COMPUTE WS-C = (WS-A + WS-B) * 2 - 1.")); }
#[test] fn compute_nested_parens() { compile_ok(&p("    COMPUTE WS-C = ((WS-A + 1) * (WS-B + 2)).")); }
#[test] fn compute_with_function() { compile_ok(&p("    COMPUTE WS-C = FUNCTION SQRT(WS-A).")); }
#[test] fn compute_mod() { compile_ok(&p("    COMPUTE WS-C = FUNCTION MOD(WS-A 3).")); }
