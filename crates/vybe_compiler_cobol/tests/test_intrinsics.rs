use vybe_parser_cobol::parse;
use vybe_compiler_cobol::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

fn p(data: &str, body: &str) -> String {
    format!("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.", data, body)
}

fn d() -> &'static str { "01 R PIC 9(10) VALUE 0.\n01 A PIC 9(10) VALUE 10.\n01 B PIC 9(10) VALUE 20.\n01 C PIC 9(10) VALUE 30." }

// ── Math functions ─────────────────────────────────────────
#[test] fn func_sqrt() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION SQRT(16).")); }
#[test] fn func_abs() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION ABS(-5).")); }
#[test] fn func_mod() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION MOD(17 5).")); }
#[test] fn func_rem() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION REM(17 5).")); }
#[test] fn func_max() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION MAX(10 20 30).")); }
#[test] fn func_min() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION MIN(10 20 30).")); }
#[test] fn func_sum() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION SUM(1 2 3 4 5).")); }
#[test] fn func_integer() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION INTEGER(3.7).")); }
#[test] fn func_power() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION POWER(2 10).")); }

// ── Trigonometric ──────────────────────────────────────────
#[test] fn func_sin() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION SIN(1).")); }
#[test] fn func_cos() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION COS(0).")); }
#[test] fn func_tan() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION TAN(1).")); }
#[test] fn func_asin() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION ASIN(1).")); }
#[test] fn func_acos() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION ACOS(0).")); }
#[test] fn func_atan() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION ATAN(1).")); }

// ── Logarithmic / exponential ──────────────────────────────
#[test] fn func_log() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION LOG(10).")); }
#[test] fn func_log10() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION LOG10(100).")); }
#[test] fn func_exp() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION EXP(1).")); }

// ── Rounding ───────────────────────────────────────────────
#[test] fn func_ceiling() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION CEILING(3.2).")); }
#[test] fn func_floor() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION FLOOR(3.7).")); }
#[test] fn func_sign() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION SIGN(-5).")); }

// ── Statistical ────────────────────────────────────────────
#[test] fn func_mean() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION MEAN(10 20 30).")); }
#[test] fn func_median() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION MEDIAN(10 20 30).")); }
#[test] fn func_variance() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION VARIANCE(10 20 30).")); }
#[test] fn func_random() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION RANDOM.")); }

// ── Date/Time ──────────────────────────────────────────────
#[test] fn func_current_date() { compile_ok(&p("01 D PIC X(21).", "    MOVE FUNCTION CURRENT-DATE TO D.")); }
#[test] fn func_when_compiled() { compile_ok(&p("01 D PIC X(21).", "    MOVE FUNCTION WHEN-COMPILED TO D.")); }

// ── Conversion ─────────────────────────────────────────────
#[test] fn func_numval() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION NUMVAL(\"12345\").")); }
#[test] fn func_ord() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION ORD(\"A\").")); }
#[test] fn func_char() { compile_ok(&p("01 C PIC X(1).", "    MOVE FUNCTION CHAR(65) TO C.")); }
#[test] fn func_test_numval() { compile_ok(&p(d(), "    COMPUTE R = FUNCTION TEST-NUMVAL(\"123\").")); }
