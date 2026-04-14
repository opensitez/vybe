use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// Fortran: Intrinsic functions — math, string, type conversion
// ═══════════════════════════════════════════════════════════

#[test]
fn abs_function() {
    let out = run_prints(r#"
program test
    print *, abs(-42)
end program test
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn sqrt_function() {
    let out = run_prints(r#"
program test
    print *, sqrt(25.0)
end program test
"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn mod_function() {
    let out = run_prints(r#"
program test
    print *, mod(17, 5)
end program test
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn min_max_functions() {
    let out = run_prints(r#"
program test
    print *, min(3, 7)
    print *, max(3, 7)
end program test
"#);
    assert_eq!(out, vec!["3", "7"]);
}

#[test]
fn trig_functions() {
    compile_ok(r#"
program test
    real :: x
    x = sin(1.0)
    x = cos(1.0)
    x = tan(1.0)
    print *, x
end program test
"#);
}

#[test]
fn exp_log_functions() {
    compile_ok(r#"
program test
    real :: x
    x = exp(1.0)
    x = log(2.718)
    print *, x
end program test
"#);
}

#[test]
fn len_trim_function() {
    let out = run_prints(r#"
program test
    character(len=20) :: s
    s = "hello"
    print *, len(s)
end program test
"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn trim_function() {
    let out = run_prints(r#"
program test
    print *, trim("  hello  ")
end program test
"#);
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn int_real_conversion() {
    compile_ok(r#"
program test
    integer :: n
    real :: x
    n = int(3.7)
    x = real(42)
    print *, n
    print *, x
end program test
"#);
}
