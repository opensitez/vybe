use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn math_sin_zero() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", sin(0.0)); return 0; }"),
        vec!["0.0"]
    );
}
#[test]
fn math_cos_zero() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", cos(0.0)); return 0; }"),
        vec!["1.0"]
    );
}
#[test]
fn math_tan_zero() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", tan(0.0)); return 0; }"),
        vec!["0.0"]
    );
}
#[test]
fn math_sin_pi_half() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { printf(\"%.1f\", sin(1.5707963267948966)); return 0; }"
        ),
        vec!["1.0"]
    );
}
#[test]
fn math_cos_pi_half() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { printf(\"%.1f\", cos(1.5707963267948966)); return 0; }"
        ),
        vec!["0.0"]
    );
}
#[test]
fn math_tan_pi_quarter() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { printf(\"%.1f\", tan(0.7853981633974483)); return 0; }"
        ),
        vec!["1.0"]
    );
}
#[test]
fn math_sin_negative() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { printf(\"%.1f\", sin(-1.5707963267948966)); return 0; }"
        ),
        vec!["-1.0"]
    );
}
#[test]
fn math_cos_negative() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { printf(\"%.1f\", cos(-1.5707963267948966)); return 0; }"
        ),
        vec!["0.0"]
    );
} // cos is even function
#[test]
fn math_tan_negative() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { printf(\"%.1f\", tan(-0.7853981633974483)); return 0; }"
        ),
        vec!["-1.0"]
    );
}
#[test]
fn math_sinf_float() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", sinf(0.0f)); return 0; }"),
        vec!["0.0"]
    );
}
#[test]
fn math_cosf_float() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", cosf(0.0f)); return 0; }"),
        vec!["1.0"]
    );
}
#[test]
fn math_tanf_float() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", tanf(0.0f)); return 0; }"),
        vec!["0.0"]
    );
}
#[test]
fn math_sin_nan() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%d\", isnan(sin(NAN))); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn math_cos_inf() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%d\", isnan(cos(INFINITY))); return 0; }"),
        vec!["1"]
    );
} // Domain error
