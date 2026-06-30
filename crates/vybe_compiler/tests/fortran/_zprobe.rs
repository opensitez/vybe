use super::helpers::{compile_ok, run_prints};

macro_rules! probe {
    ($name:expr, $src:expr) => {
        eprintln!("=== {} ===", $name);
        eprintln!("  parse_ok: {}", super::helpers::parse_ok($src));
        match std::panic::catch_unwind(|| compile_ok($src)) {
            Ok(_) => eprintln!("  compile: ok"),
            Err(_) => eprintln!("  compile: panic"),
        }
        match std::panic::catch_unwind(|| run_prints($src)) {
            Ok(out) => eprintln!("  run: {:?}", out),
            Err(_) => eprintln!("  run: panic"),
        }
    };
}

#[test]
fn zprobe() {
    probe!("goto", "program t\ninteger :: x = 0\ngoto 10\nx = 999\n10 continue\nprint *, x\nend program t\n");
    probe!("arith", "program t\nreal :: x = -1.0\nif (x) 10, 20, 30\n10 print *, 'neg'; goto 99\n20 print *, 'zer'; goto 99\n30 print *, 'pos'\n99 continue\nend program t\n");
    probe!("data", "program t\ninteger :: x, y\ndata x /42/, y /99/\nprint *, x + y\nend program t\n");
    probe!("equiv", "program t\ninteger :: a, b\nequivalence (a, b)\na = 42\nprint *, b\nend program t\n");
    probe!("assign_goto", "program t\ninteger :: lab\nassign 10 to lab\ngoto lab\nprint *, 'skip'\n10 continue\nprint *, 'ok'\nend program t\n");
    probe!("computed", "program t\ninteger :: n = 2\ngo to (10, 20, 30), n\n10 print *, 'one'; goto 99\n20 print *, 'two'; goto 99\n30 print *, 'three'\n99 continue\nend program t\n");
    probe!("stmtfn", "program t\nreal :: x\nsq(x) = x * x\nprint *, sq(4.0)\nend program t\n");
    probe!("external", "program t\nexternal :: dbl\nprint *, dbl(3.0)\ncontains\nfunction dbl(x)\nreal :: dbl, x\ndbl = x * 2.0\nend function dbl\nend program t\n");
    probe!("intrinsic", "program t\nintrinsic :: abs\nprint *, abs(-5)\nend program t\n");
}
