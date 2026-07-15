use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn strlcpy_basic() {
    assert_eq!(
        run_c(
            "/* BSD/macOS */\n#include <string.h>\nint main() { char dest[10]; size_t res = strlcpy(dest, \"hello\", sizeof(dest)); printf(\"%s %d\", dest, (int)res); return 0; }"
        ),
        vec!["hello 5"]
    );
}
#[test]
fn strlcpy_truncation() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char dest[4]; size_t res = strlcpy(dest, \"hello\", sizeof(dest)); printf(\"%s %d\", dest, (int)res); return 0; }"
        ),
        vec!["hel 5"]
    );
}
#[test]
fn strlcpy_zero_size() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char dest[4] = \"abc\"; size_t res = strlcpy(dest, \"hello\", 0); printf(\"%s %d\", dest, (int)res); return 0; }"
        ),
        vec!["abc 5"]
    );
}
#[test]
fn strlcat_basic() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char dest[10] = \"hi \"; size_t res = strlcat(dest, \"there\", sizeof(dest)); printf(\"%s %d\", dest, (int)res); return 0; }"
        ),
        vec!["hi there 8"]
    );
}
#[test]
fn strlcat_truncation() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char dest[8] = \"hello \"; size_t res = strlcat(dest, \"world\", sizeof(dest)); printf(\"%s %d\", dest, (int)res); return 0; }"
        ),
        vec!["hello w 11"]
    );
} // len("hello ") + len("world")
#[test]
fn strlcat_no_null_in_dest() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char dest[4] = \"abcd\"; size_t res = strlcat(dest, \"e\", sizeof(dest)); printf(\"%d %c%c%c%c\", (int)res, dest[0], dest[1], dest[2], dest[3]); return 0; }"
        ),
        vec!["5 abcd"]
    );
} // returns size + len(src), dest unchanged
#[test]
fn stpncpy_basic() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char dest[10]; char *p = stpncpy(dest, \"hello\", 10); printf(\"%d %s\", (int)(p - dest), dest); return 0; }"
        ),
        vec!["5 hello"]
    );
}
#[test]
fn stpncpy_truncation() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char dest[4]; char *p = stpncpy(dest, \"hello\", 3); dest[3] = '\\0'; printf(\"%d %s\", (int)(p - dest), dest); return 0; }"
        ),
        vec!["3 hel"]
    );
}
#[test]
fn stpcpy_basic() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char dest[10]; char *p = stpcpy(dest, \"hello\"); printf(\"%d %s\", (int)(p - dest), dest); return 0; }"
        ),
        vec!["5 hello"]
    );
}
#[test]
fn strdup_basic() {
    assert_eq!(
        run_c(
            "#include <string.h>\n#include <stdlib.h>\nint main() { char *s = strdup(\"hello\"); printf(\"%s\", s); free(s); return 0; }"
        ),
        vec!["hello"]
    );
}
#[test]
fn strndup_basic() {
    assert_eq!(
        run_c(
            "#include <string.h>\n#include <stdlib.h>\nint main() { char *s = strndup(\"hello\", 3); printf(\"%s\", s); free(s); return 0; }"
        ),
        vec!["hel"]
    );
}
#[test]
fn strndup_long() {
    assert_eq!(
        run_c(
            "#include <string.h>\n#include <stdlib.h>\nint main() { char *s = strndup(\"hi\", 10); printf(\"%s\", s); free(s); return 0; }"
        ),
        vec!["hi"]
    );
}
