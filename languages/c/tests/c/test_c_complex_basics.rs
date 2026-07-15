use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn complex_basic_creation() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z = 1.0 + 2.0 * I; printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn complex_real_part() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z = 1.5 + 2.5 * I; printf(\"%.1f\", creal(z)); return 0; }"
        ),
        vec!["1.5"]
    );
}
#[test]
fn complex_imag_part() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z = 1.5 + 2.5 * I; printf(\"%.1f\", cimag(z)); return 0; }"
        ),
        vec!["2.5"]
    );
}
#[test]
fn complex_addition() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z1 = 1.0 + 2.0 * I; double complex z2 = 3.0 + 4.0 * I; double complex z3 = z1 + z2; printf(\"%.1f %.1f\", creal(z3), cimag(z3)); return 0; }"
        ),
        vec!["4.0 6.0"]
    );
}
#[test]
fn complex_subtraction() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z1 = 5.0 + 6.0 * I; double complex z2 = 2.0 + 3.0 * I; double complex z3 = z1 - z2; printf(\"%.1f %.1f\", creal(z3), cimag(z3)); return 0; }"
        ),
        vec!["3.0 3.0"]
    );
}
#[test]
fn complex_multiplication() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z1 = 1.0 + 2.0 * I; double complex z2 = 2.0 + 3.0 * I; double complex z3 = z1 * z2; /* (1+2i)*(2+3i) = 2 + 3i + 4i - 6 = -4 + 7i */ printf(\"%.1f %.1f\", creal(z3), cimag(z3)); return 0; }"
        ),
        vec!["-4.0 7.0"]
    );
}
#[test]
fn complex_division() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z1 = -4.0 + 7.0 * I; double complex z2 = 2.0 + 3.0 * I; double complex z3 = z1 / z2; printf(\"%.1f %.1f\", creal(z3), cimag(z3)); return 0; }"
        ),
        vec!["1.0 2.0"]
    );
}
#[test]
fn complex_conjugate() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z = 1.5 + 2.5 * I; double complex z2 = conj(z); printf(\"%.1f %.1f\", creal(z2), cimag(z2)); return 0; }"
        ),
        vec!["1.5 -2.5"]
    );
}
#[test]
fn complex_magnitude() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z = 3.0 + 4.0 * I; printf(\"%.1f\", cabs(z)); return 0; }"
        ),
        vec!["5.0"]
    );
} // sqrt(3^2 + 4^2) = 5
#[test]
fn complex_phase() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z = 0.0 + 1.0 * I; printf(\"%.5f\", carg(z)); return 0; }"
        ),
        vec!["1.57080"]
    );
} // Pi/2
#[test]
fn complex_float_type() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { float complex z = 1.0f + 2.0f * I; printf(\"%.1f\", crealf(z)); return 0; }"
        ),
        vec!["1.0"]
    );
}
#[test]
fn complex_long_double_type() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { long double complex z = 1.0L + 2.0L * I; printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn complex_cmp_equality() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z1 = 1.0 + 2.0 * I; double complex z2 = 1.0 + 2.0 * I; printf(\"%d\", z1 == z2); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn complex_cmp_inequality() {
    assert_eq!(
        run_c(
            "#include <complex.h>\nint main() { double complex z1 = 1.0 + 2.0 * I; double complex z2 = 1.0 + 3.0 * I; printf(\"%d\", z1 != z2); return 0; }"
        ),
        vec!["1"]
    );
}
