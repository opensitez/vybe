use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn float_nan_macro() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { double n = NAN; printf(\"%d\", isnan(n)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn float_nan_not_equal_to_self() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { double n = NAN; printf(\"%d\", n != n); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn float_nan_addition() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { double n = NAN + 5.0; printf(\"%d\", isnan(n)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn float_nan_multiplication() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { double n = NAN * 2.0; printf(\"%d\", isnan(n)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn float_nan_zero_div_zero() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { double n = 0.0 / 0.0; printf(\"%d\", isnan(n)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn float_nan_inf_minus_inf() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { double n = INFINITY - INFINITY; printf(\"%d\", isnan(n)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn float_nan_inf_div_inf() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { double n = INFINITY / INFINITY; printf(\"%d\", isnan(n)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn float_nan_zero_times_inf() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { double n = 0.0 * INFINITY; printf(\"%d\", isnan(n)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn float_nan_fmod() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { double n = fmod(5.0, 0.0); printf(\"%d\", isnan(n)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn float_nan_sqrt_negative() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { double n = sqrt(-1.0); printf(\"%d\", isnan(n)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn float_nan_log_negative() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { double n = log(-1.0); printf(\"%d\", isnan(n)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn float_nan_asin_out_of_range() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { double n = asin(2.0); printf(\"%d\", isnan(n)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn float_nan_pow_one_inf() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { double n = pow(1.0, INFINITY); printf(\"%d\", n == 1.0); return 0; }"
        ),
        vec!["1"]
    );
} // IEEE says 1.0, not NaN
#[test]
fn float_nan_comparison_less() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { double n = NAN; printf(\"%d\", !(n < 0.0) && !(n > 0.0) && !(n == 0.0)); return 0; }"
        ),
        vec!["1"]
    );
} // all false
