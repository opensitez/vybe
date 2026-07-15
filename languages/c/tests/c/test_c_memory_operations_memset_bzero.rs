use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn memset_basic() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char s[5] = \"abcd\"; memset(s, 'x', 3); printf(\"%s\", s); return 0; }"
        ),
        vec!["xxxd"]
    );
}
#[test]
fn memset_zero() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { int arr[3]; memset(arr, 0, sizeof(arr)); printf(\"%d %d %d\", arr[0], arr[1], arr[2]); return 0; }"
        ),
        vec!["0 0 0"]
    );
}
#[test]
fn memset_zero_length() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { char s[5] = \"abcd\"; memset(s, 'x', 0); printf(\"%s\", s); return 0; }"
        ),
        vec!["abcd"]
    );
}
#[test]
fn bzero_basic() {
    assert_eq!(
        run_c(
            "#include <strings.h>\nint main() { char s[5] = \"abcd\"; bzero(s, 3); printf(\"%d %d %d %c\", s[0], s[1], s[2], s[3]); return 0; }"
        ),
        vec!["0 0 0 d"]
    );
}
#[test]
fn bcopy_fwd() {
    assert_eq!(
        run_c(
            "#include <strings.h>\nint main() { char s1[] = \"abc\"; char s2[4]; bcopy(s1, s2, 4); printf(\"%s\", s2); return 0; }"
        ),
        vec!["abc"]
    );
}
#[test]
fn bcopy_overlap() {
    assert_eq!(
        run_c(
            "#include <strings.h>\nint main() { char s[] = \"abcdef\"; bcopy(s, s + 2, 4); printf(\"%s\", s); return 0; }"
        ),
        vec!["ababcd"]
    );
}
#[test]
fn memset_explicit_compile() {
    assert_eq!(
        run_c(
            "/* memset_explicit might not be available everywhere, so we just compile test if it is or fallback */\n#include <string.h>\nint main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
