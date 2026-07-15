use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn fenv_divzero_flag() {
    assert_eq!(
        run_c(
            "#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { feclearexcept(FE_ALL_EXCEPT); double x = 1.0 / 0.0; printf(\"%d\", (fetestexcept(FE_DIVBYZERO) != 0)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn fenv_invalid_flag() {
    assert_eq!(
        run_c(
            "#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { feclearexcept(FE_ALL_EXCEPT); double x = 0.0 / 0.0; printf(\"%d\", (fetestexcept(FE_INVALID) != 0)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn fenv_inexact_flag() {
    assert_eq!(
        run_c(
            "#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { feclearexcept(FE_ALL_EXCEPT); double x = 1.0 / 3.0; printf(\"%d\", (fetestexcept(FE_INEXACT) != 0)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn fenv_overflow_flag() {
    assert_eq!(
        run_c(
            "#include <fenv.h>\n#include <float.h>\n#pragma STDC FENV_ACCESS ON\nint main() { feclearexcept(FE_ALL_EXCEPT); double x = DBL_MAX * 2.0; printf(\"%d\", (fetestexcept(FE_OVERFLOW) != 0)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn fenv_underflow_flag() {
    assert_eq!(
        run_c(
            "#include <fenv.h>\n#include <float.h>\n#pragma STDC FENV_ACCESS ON\nint main() { feclearexcept(FE_ALL_EXCEPT); double x = DBL_MIN / 1e100; printf(\"%d\", (fetestexcept(FE_UNDERFLOW) != 0)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn fenv_clear_exceptions() {
    assert_eq!(
        run_c(
            "#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { feclearexcept(FE_ALL_EXCEPT); double x = 1.0 / 0.0; feclearexcept(FE_ALL_EXCEPT); printf(\"%d\", (fetestexcept(FE_ALL_EXCEPT) == 0)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn fenv_raise_exception() {
    assert_eq!(
        run_c(
            "#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { feclearexcept(FE_ALL_EXCEPT); feraiseexcept(FE_DIVBYZERO); printf(\"%d\", (fetestexcept(FE_DIVBYZERO) != 0)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn fenv_save_restore_env() {
    assert_eq!(
        run_c(
            "#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { fenv_t env; fegetenv(&env); feraiseexcept(FE_INVALID); fesetenv(&env); printf(\"%d\", fetestexcept(FE_INVALID) == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn fenv_hold_update_env() {
    assert_eq!(
        run_c(
            "#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { fenv_t env; feholdexcept(&env); double x = 1.0 / 0.0; feupdateenv(&env); printf(\"%d\", (fetestexcept(FE_DIVBYZERO) != 0)); return 0; }"
        ),
        vec!["1"]
    );
} // feupdateenv applies currently raised on top of saved
#[test]
fn fenv_multiple_exceptions() {
    assert_eq!(
        run_c(
            "#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { feclearexcept(FE_ALL_EXCEPT); feraiseexcept(FE_DIVBYZERO | FE_INVALID); printf(\"%d\", (fetestexcept(FE_DIVBYZERO | FE_INVALID) == (FE_DIVBYZERO | FE_INVALID))); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn fenv_test_specific_exception() {
    assert_eq!(
        run_c(
            "#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { feclearexcept(FE_ALL_EXCEPT); feraiseexcept(FE_DIVBYZERO); printf(\"%d\", fetestexcept(FE_DIVBYZERO | FE_INVALID) == FE_DIVBYZERO); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn fenv_fe_dfl_env() {
    assert_eq!(
        run_c(
            "#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { feraiseexcept(FE_INVALID); fesetenv(FE_DFL_ENV); printf(\"%d\", fetestexcept(FE_ALL_EXCEPT) == 0); return 0; }"
        ),
        vec!["1"]
    );
} // Default environment clears all exceptions
#[test]
fn fenv_no_exception_on_normal_ops() {
    assert_eq!(
        run_c(
            "#include <fenv.h>\n#pragma STDC FENV_ACCESS ON\nint main() { feclearexcept(FE_ALL_EXCEPT); double x = 2.0 + 2.0; printf(\"%d\", fetestexcept(FE_ALL_EXCEPT) == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn fenv_sqrt_negative_is_invalid() {
    assert_eq!(
        run_c(
            "#include <fenv.h>\n#include <math.h>\n#pragma STDC FENV_ACCESS ON\nint main() { feclearexcept(FE_ALL_EXCEPT); double x = sqrt(-1.0); printf(\"%d\", (fetestexcept(FE_INVALID) != 0)); return 0; }"
        ),
        vec!["1"]
    );
}
