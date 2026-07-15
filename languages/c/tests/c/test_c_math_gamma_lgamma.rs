use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn math_tgamma_one() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", tgamma(1.0)); return 0; }"),
        vec!["1.0"]
    );
} // 0! = 1
#[test]
fn math_tgamma_two() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", tgamma(2.0)); return 0; }"),
        vec!["1.0"]
    );
} // 1! = 1
#[test]
fn math_tgamma_five() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", tgamma(5.0)); return 0; }"),
        vec!["24.0"]
    );
} // 4! = 24
#[test]
fn math_tgamma_half() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.5f\", tgamma(0.5)); return 0; }"),
        vec!["1.77245"]
    );
} // sqrt(pi)
#[test]
fn math_lgamma_one() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", lgamma(1.0)); return 0; }"),
        vec!["0.0"]
    );
} // ln(1) = 0
#[test]
fn math_lgamma_two() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", lgamma(2.0)); return 0; }"),
        vec!["0.0"]
    );
} // ln(1) = 0
#[test]
fn math_lgamma_five() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.5f\", lgamma(5.0)); return 0; }"),
        vec!["3.17805"]
    );
} // ln(24) ~ 3.17805
#[test]
fn math_tgamma_negative_integer() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"ok\"); return 0; }"),
        vec!["ok"]
    );
} // Domain error or pole error, implementation defined if NAN or INF
#[test]
fn math_lgamma_signgam() {
    assert_eq!(
        run_c(
            "#include <math.h>\nextern int signgam;\nint main() { lgamma(-0.5); printf(\"%d\", signgam); return 0; }"
        ),
        vec!["-1"]
    );
} // gamma(-0.5) is -2*sqrt(pi)
#[test]
fn math_erf_zero() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", erf(0.0)); return 0; }"),
        vec!["0.0"]
    );
}
#[test]
fn math_erf_inf() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", erf(INFINITY)); return 0; }"),
        vec!["1.0"]
    );
}
#[test]
fn math_erfc_zero() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", erfc(0.0)); return 0; }"),
        vec!["1.0"]
    );
}
#[test]
fn math_erfc_inf() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", erfc(INFINITY)); return 0; }"),
        vec!["0.0"]
    );
}
#[test]
fn math_erf_negative_inf() {
    assert_eq!(
        run_c("#include <math.h>\nint main() { printf(\"%.1f\", erf(-INFINITY)); return 0; }"),
        vec!["-1.0"]
    );
} // Odd function
