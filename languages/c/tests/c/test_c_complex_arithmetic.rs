use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn complex_pow() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z = cpow(2.0 + 0.0 * I, 3.0 + 0.0 * I); printf(\"%.1f\", creal(z)); return 0; }"
        ),
        vec!["8.0"]
    );
}
#[test]
fn complex_sqrt() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z = csqrt(-1.0 + 0.0 * I); printf(\"%.1f\", cimag(z)); return 0; }"
        ),
        vec!["1.0"]
    );
}
#[test]
fn complex_exp() {
    assert_eq!(
        run_c(
            "#include <complex.h>\n#include <math.h>\nint main() { double complex z = cexp(0.0 + M_PI * I); printf(\"%.1f\", creal(z)); return 0; }"
        ),
        vec!["-1.0"]
    );
} // Euler's identity: e^(i*pi) = -1
#[test]
fn complex_log() {
    assert_eq!(
        run_c(
            "#include <complex.h>\n#include <math.h>\nint main() { double complex z = clog(-1.0 + 0.0 * I); printf(\"%.5f\", cimag(z)); return 0; }"
        ),
        vec!["3.14159"]
    );
} // ln(-1) = i*pi
#[test]
fn complex_sin() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z = csin(0.0 + 1.0 * I); printf(\"%.5f\", cimag(z)); return 0; }"
        ),
        vec!["1.17520"]
    );
} // sin(i) = i*sinh(1)
#[test]
fn complex_cos() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z = ccos(0.0 + 1.0 * I); printf(\"%.5f\", creal(z)); return 0; }"
        ),
        vec!["1.54308"]
    );
} // cos(i) = cosh(1)
#[test]
fn complex_tan() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z = ctan(0.0 + 1.0 * I); printf(\"%.5f\", cimag(z)); return 0; }"
        ),
        vec!["0.76159"]
    );
} // tan(i) = i*tanh(1)
#[test]
fn complex_asin() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z = casin(2.0 + 0.0 * I); printf(\"%.5f\", creal(z)); return 0; }"
        ),
        vec!["1.57080"]
    );
}
#[test]
fn complex_acos() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z = cacos(2.0 + 0.0 * I); printf(\"%.5f\", cimag(z)); return 0; }"
        ),
        vec!["-1.31696"]
    );
} // acosh(2)
#[test]
fn complex_atan() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z = catan(0.0 + 2.0 * I); printf(\"%.5f\", cimag(z)); return 0; }"
        ),
        vec!["0.54931"]
    );
} // atanh(2) ? No wait, let's test compile
#[test]
fn complex_sinh() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z = csinh(0.0 + 1.0 * I); printf(\"%.5f\", cimag(z)); return 0; }"
        ),
        vec!["0.84147"]
    );
} // sinh(i) = i*sin(1)
#[test]
fn complex_cosh() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z = ccosh(0.0 + 1.0 * I); printf(\"%.5f\", creal(z)); return 0; }"
        ),
        vec!["0.54030"]
    );
} // cosh(i) = cos(1)
#[test]
fn complex_tanh() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z = ctanh(0.0 + 1.0 * I); printf(\"%.5f\", cimag(z)); return 0; }"
        ),
        vec!["1.55741"]
    );
} // tanh(i) = i*tan(1)
#[test]
fn complex_proj() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z = cproj(1.0 + INFINITY * I); printf(\"%d\", isinf(creal(z))); return 0; }"
        ),
        vec!["1"]
    );
} // projects to infinity on Riemann sphere
