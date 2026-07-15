use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn tgmath_sin_double() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { double x = 0.0; printf(\"%.1f\", sin(x)); return 0; }"
        ),
        vec!["0.0"]
    );
}
#[test]
fn tgmath_sin_float() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { float x = 0.0f; printf(\"%.1f\", sin(x)); return 0; }"
        ),
        vec!["0.0"]
    );
}
#[test]
fn tgmath_sin_long_double() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { long double x = 0.0L; printf(\"%.1f\", (double)sin(x)); return 0; }"
        ),
        vec!["0.0"]
    );
}
#[test]
fn tgmath_cos_double() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { double x = 0.0; printf(\"%.1f\", cos(x)); return 0; }"
        ),
        vec!["1.0"]
    );
}
#[test]
fn tgmath_pow_float() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { float x = 2.0f, y = 3.0f; printf(\"%.1f\", pow(x, y)); return 0; }"
        ),
        vec!["8.0"]
    );
}
#[test]
fn tgmath_pow_mixed() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { double x = 2.0; float y = 3.0f; printf(\"%.1f\", pow(x, y)); return 0; }"
        ),
        vec!["8.0"]
    );
} // Promotes to double
#[test]
fn tgmath_sqrt_int() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { int x = 9; printf(\"%.1f\", sqrt(x)); return 0; }"
        ),
        vec!["3.0"]
    );
} // Int promotes to double
#[test]
fn tgmath_fabs_float() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { float x = -5.0f; printf(\"%.1f\", fabs(x)); return 0; }"
        ),
        vec!["5.0"]
    );
}
#[test]
fn tgmath_exp_double() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { double x = 0.0; printf(\"%.1f\", exp(x)); return 0; }"
        ),
        vec!["1.0"]
    );
}
#[test]
fn tgmath_log_double() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { double x = 1.0; printf(\"%.1f\", log(x)); return 0; }"
        ),
        vec!["0.0"]
    );
}
#[test]
fn tgmath_fmod_float() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { float x = 5.0f, y = 2.0f; printf(\"%.1f\", fmod(x, y)); return 0; }"
        ),
        vec!["1.0"]
    );
}
#[test]
fn tgmath_atan2_mixed() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { float y = 0.0f; double x = 1.0; printf(\"%.1f\", atan2(y, x)); return 0; }"
        ),
        vec!["0.0"]
    );
}
#[test]
fn tgmath_cbrt_int() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { int x = 8; printf(\"%.1f\", cbrt(x)); return 0; }"
        ),
        vec!["2.0"]
    );
}
#[test]
fn tgmath_ceil_float() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { float x = 2.3f; printf(\"%.1f\", ceil(x)); return 0; }"
        ),
        vec!["3.0"]
    );
}
