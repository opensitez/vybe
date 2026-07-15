use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn math_fdim_positive_diff() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", fdim(5.0, 3.0)); return 0; }"),
        vec!["2.0"]
    );
}
#[test]
fn math_fdim_negative_diff() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", fdim(3.0, 5.0)); return 0; }"),
        vec!["0.0"]
    );
} // Difference is negative, returns 0.0
#[test]
fn math_fdim_equal() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", fdim(5.0, 5.0)); return 0; }"),
        vec!["0.0"]
    );
}
#[test]
fn math_fmax_first() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", fmax(5.0, 3.0)); return 0; }"),
        vec!["5.0"]
    );
}
#[test]
fn math_fmax_second() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", fmax(3.0, 5.0)); return 0; }"),
        vec!["5.0"]
    );
}
#[test]
fn math_fmin_first() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", fmin(3.0, 5.0)); return 0; }"),
        vec!["3.0"]
    );
}
#[test]
fn math_fmin_second() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", fmin(5.0, 3.0)); return 0; }"),
        vec!["3.0"]
    );
}
#[test]
fn math_fmax_nan_first() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", fmax(NAN, 3.0)); return 0; }"),
        vec!["3.0"]
    );
} // fmax ignores NaN if possible
#[test]
fn math_fmax_nan_second() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", fmax(5.0, NAN)); return 0; }"),
        vec!["5.0"]
    );
}
#[test]
fn math_fmin_nan_first() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", fmin(NAN, 3.0)); return 0; }"),
        vec!["3.0"]
    );
}
#[test]
fn math_fmin_nan_second() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", fmin(5.0, NAN)); return 0; }"),
        vec!["5.0"]
    );
}
#[test]
fn math_fmax_nan_both() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%d\", isnan(fmax(NAN, NAN))); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn math_fmin_nan_both() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%d\", isnan(fmin(NAN, NAN))); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn math_fdim_nan() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%d\", isnan(fdim(5.0, NAN))); return 0; }"),
        vec!["1"]
    );
}
