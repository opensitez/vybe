use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn strtol_basic() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { char *end; printf(\"%ld\", strtol(\"123\", &end, 10)); return 0; }"
        ),
        vec!["123"]
    );
}
#[test]
fn strtol_hex() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { char *end; printf(\"%ld\", strtol(\"0x1A\", &end, 16)); return 0; }"
        ),
        vec!["26"]
    );
}
#[test]
fn strtol_base_0_hex() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { char *end; printf(\"%ld\", strtol(\"0x1A\", &end, 0)); return 0; }"
        ),
        vec!["26"]
    );
}
#[test]
fn strtol_base_0_octal() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { char *end; printf(\"%ld\", strtol(\"012\", &end, 0)); return 0; }"
        ),
        vec!["10"]
    );
}
#[test]
fn strtol_negative() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { char *end; printf(\"%ld\", strtol(\"-123\", &end, 10)); return 0; }"
        ),
        vec!["-123"]
    );
}
#[test]
fn strtol_with_endptr() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { char *end; strtol(\"123abc\", &end, 10); printf(\"%s\", end); return 0; }"
        ),
        vec!["abc"]
    );
}
#[test]
fn strtoul_basic() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { char *end; printf(\"%lu\", strtoul(\"4294967295\", &end, 10)); return 0; }"
        ),
        vec!["4294967295"]
    );
}
#[test]
fn strtoll_basic() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { char *end; printf(\"%lld\", strtoll(\"123456789012345\", &end, 10)); return 0; }"
        ),
        vec!["123456789012345"]
    );
}
#[test]
fn strtod_basic() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { char *end; printf(\"%.2f\", strtod(\"3.14\", &end)); return 0; }"
        ),
        vec!["3.14"]
    );
}
#[test]
fn strtod_scientific() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { char *end; printf(\"%.2f\", strtod(\"1.2e3\", &end)); return 0; }"
        ),
        vec!["1200.00"]
    );
}
#[test]
fn strtod_hex_float() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { char *end; printf(\"%.2f\", strtod(\"0x1.8p1\", &end)); return 0; }"
        ),
        vec!["3.00"]
    );
}
#[test]
fn strtod_inf() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\n#include <math.h>\nint main() { char *end; printf(\"%d\", isinf(strtod(\"INF\", &end))); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn strtod_nan() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\n#include <math.h>\nint main() { char *end; printf(\"%d\", isnan(strtod(\"NaN\", &end))); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn atoi_basic() {
    assert_eq!(
        run_c("#include <stdlib.h>\nint main() { printf(\"%d\", atoi(\"123\")); return 0; }"),
        vec!["123"]
    );
}
#[test]
fn atof_basic() {
    assert_eq!(
        run_c("#include <stdlib.h>\nint main() { printf(\"%.2f\", atof(\"3.14\")); return 0; }"),
        vec!["3.14"]
    );
}
