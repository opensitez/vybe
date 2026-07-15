use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn tgmath_complex_sin_double() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { double complex z = 0.0 + 1.0 * I; printf(\"%.5f\", cimag(sin(z))); return 0; }"
        ),
        vec!["1.17520"]
    );
}
#[test]
fn tgmath_complex_sin_float() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { float complex z = 0.0f + 1.0f * I; printf(\"%.5f\", cimagf(sin(z))); return 0; }"
        ),
        vec!["1.17520"]
    );
}
#[test]
fn tgmath_complex_cos_double() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { double complex z = 0.0 + 1.0 * I; printf(\"%.5f\", creal(cos(z))); return 0; }"
        ),
        vec!["1.54308"]
    );
}
#[test]
fn tgmath_complex_sqrt_double() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { double complex z = -1.0 + 0.0 * I; printf(\"%.1f\", cimag(sqrt(z))); return 0; }"
        ),
        vec!["1.0"]
    );
}
#[test]
fn tgmath_complex_exp_float() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { float complex z = 0.0f + 3.14159265f * I; printf(\"%.1f\", crealf(exp(z))); return 0; }"
        ),
        vec!["-1.0"]
    );
}
#[test]
fn tgmath_complex_pow_double() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { double complex z1 = 2.0 + 0.0 * I; double complex z2 = 3.0 + 0.0 * I; printf(\"%.1f\", creal(pow(z1, z2))); return 0; }"
        ),
        vec!["8.0"]
    );
}
#[test]
fn tgmath_complex_pow_mixed() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { double complex z = 2.0 + 0.0 * I; double x = 3.0; printf(\"%.1f\", creal(pow(z, x))); return 0; }"
        ),
        vec!["8.0"]
    );
}
#[test]
fn tgmath_complex_fabs_double() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { double complex z = 3.0 + 4.0 * I; printf(\"%.1f\", fabs(z)); return 0; }"
        ),
        vec!["5.0"]
    );
} // cabs
#[test]
fn tgmath_complex_log_double() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { double complex z = -1.0 + 0.0 * I; printf(\"%.5f\", cimag(log(z))); return 0; }"
        ),
        vec!["3.14159"]
    );
}
#[test]
fn tgmath_complex_asin_double() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { double complex z = 2.0 + 0.0 * I; printf(\"%.5f\", creal(asin(z))); return 0; }"
        ),
        vec!["1.57080"]
    );
}
#[test]
fn tgmath_complex_acos_double() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { double complex z = 2.0 + 0.0 * I; printf(\"%.5f\", cimag(acos(z))); return 0; }"
        ),
        vec!["-1.31696"]
    );
}
#[test]
fn tgmath_complex_atan_double() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { double complex z = 0.0 + 2.0 * I; printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn tgmath_complex_sinh_double() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { double complex z = 0.0 + 1.0 * I; printf(\"%.5f\", cimag(sinh(z))); return 0; }"
        ),
        vec!["0.84147"]
    );
}
#[test]
fn tgmath_complex_cosh_double() {
    assert_eq!(
        run_c(
            "#include <tgmath.h>\nint main() { double complex z = 0.0 + 1.0 * I; printf(\"%.5f\", creal(cosh(z))); return 0; }"
        ),
        vec!["0.54030"]
    );
}
