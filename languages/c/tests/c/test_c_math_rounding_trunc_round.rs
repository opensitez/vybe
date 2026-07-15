use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn math_trunc_positive() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", trunc(2.7)); return 0; }"),
        vec!["2.0"]
    );
}
#[test]
fn math_trunc_negative() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", trunc(-2.7)); return 0; }"),
        vec!["-2.0"]
    );
}
#[test]
fn math_round_positive_down() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", round(2.3)); return 0; }"),
        vec!["2.0"]
    );
}
#[test]
fn math_round_positive_up() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", round(2.8)); return 0; }"),
        vec!["3.0"]
    );
}
#[test]
fn math_round_positive_half() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", round(2.5)); return 0; }"),
        vec!["3.0"]
    );
} // Halfway cases rounded away from zero
#[test]
fn math_round_negative_down() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", round(-2.3)); return 0; }"),
        vec!["-2.0"]
    );
}
#[test]
fn math_round_negative_up() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", round(-2.8)); return 0; }"),
        vec!["-3.0"]
    );
}
#[test]
fn math_round_negative_half() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", round(-2.5)); return 0; }"),
        vec!["-3.0"]
    );
} // Halfway cases rounded away from zero
#[test]
fn math_lround_positive() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%ld\", lround(2.5)); return 0; }"),
        vec!["3"]
    );
}
#[test]
fn math_llround_negative() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%lld\", llround(-2.5)); return 0; }"),
        vec!["-3"]
    );
}
#[test]
fn math_nearbyint_positive() {
    assert_eq!(
        run_c(
            "#include <math.h>\n#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { fesetround(FE_TONEAREST); printf(\"%.1f\", nearbyint(2.5)); return 0; }"
        ),
        vec!["2.0"]
    );
} // round to even for halfway
#[test]
fn math_rint_positive() {
    assert_eq!(
        run_c(
            "#include <math.h>\n#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { fesetround(FE_UPWARD); printf(\"%.1f\", rint(2.1)); return 0; }"
        ),
        vec!["3.0"]
    );
}
#[test]
fn math_lrint_positive() {
    assert_eq!(
        run_c(
            "#include <math.h>\n#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { fesetround(FE_DOWNWARD); printf(\"%ld\", lrint(2.9)); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn math_trunc_inf() {
    assert_eq!(
        run_c(
            "#include <math.h>\nint main() { printf(\"%d\", isinf(trunc(INFINITY))); return 0; }"
        ),
        vec!["1"]
    );
}
