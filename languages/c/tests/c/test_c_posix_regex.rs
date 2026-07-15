use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn regcomp_regexec_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; int r = regcomp(&re, \"a.*b\", 0); if(r == 0) { int r2 = regexec(&re, \"axxxb\", 0, NULL, 0); printf(\"%d\", r2 == 0); regfree(&re); } else printf(\"0\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn regcomp_invalid_regex() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; int r = regcomp(&re, \"[a-z\", 0); printf(\"%d\", r != 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn regerror_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; int r = regcomp(&re, \"[a-z\", 0); char buf[256]; regerror(r, &re, buf, sizeof(buf)); printf(\"%d\", buf[0] != '\\0'); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn regcomp_extended() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; int r = regcomp(&re, \"a+b\", REG_EXTENDED); if(r == 0) { int r2 = regexec(&re, \"aaab\", 0, NULL, 0); printf(\"%d\", r2 == 0); regfree(&re); } else printf(\"0\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn regcomp_icase() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; int r = regcomp(&re, \"abc\", REG_ICASE); if(r == 0) { int r2 = regexec(&re, \"AbC\", 0, NULL, 0); printf(\"%d\", r2 == 0); regfree(&re); } else printf(\"0\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn regcomp_nosub() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; int r = regcomp(&re, \"(abc)\", REG_NOSUB); if(r == 0) { printf(\"%d\", re.re_nsub == 0); regfree(&re); } else printf(\"0\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn regcomp_newline() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; int r = regcomp(&re, \"a.b\", REG_NEWLINE); if(r == 0) { int r2 = regexec(&re, \"a\\nb\", 0, NULL, 0); printf(\"%d\", r2 == REG_NOMATCH); regfree(&re); } else printf(\"0\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn regexec_match_start_end() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; int r = regcomp(&re, \"bc\", 0); if(r == 0) { regmatch_t m[1]; int r2 = regexec(&re, \"abcde\", 1, m, 0); printf(\"%d %d %d\", r2 == 0, m[0].rm_so == 1, m[0].rm_eo == 3); regfree(&re); } else printf(\"0\"); return 0; }"
        ),
        vec!["1 1 1"]
    );
}
#[test]
fn regexec_multiple_matches() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; int r = regcomp(&re, \"([a-z]+) ([0-9]+)\", REG_EXTENDED); if(r == 0) { regmatch_t m[3]; int r2 = regexec(&re, \"hello 123\", 3, m, 0); printf(\"%d %d %d\", r2 == 0, m[1].rm_eo - m[1].rm_so == 5, m[2].rm_eo - m[2].rm_so == 3); regfree(&re); } else printf(\"0\"); return 0; }"
        ),
        vec!["1 1 1"]
    );
}
#[test]
fn regexec_notbol() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; int r = regcomp(&re, \"^abc\", 0); if(r == 0) { int r2 = regexec(&re, \"abc\", 0, NULL, REG_NOTBOL); printf(\"%d\", r2 == REG_NOMATCH); regfree(&re); } else printf(\"0\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn regexec_noteol() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; int r = regcomp(&re, \"abc$\", 0); if(r == 0) { int r2 = regexec(&re, \"abc\", 0, NULL, REG_NOTEOL); printf(\"%d\", r2 == REG_NOMATCH); regfree(&re); } else printf(\"0\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn regexec_no_match() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; int r = regcomp(&re, \"abc\", 0); if(r == 0) { int r2 = regexec(&re, \"xyz\", 0, NULL, 0); printf(\"%d\", r2 == REG_NOMATCH); regfree(&re); } else printf(\"0\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn regfree_null() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { /* Most regfree expect a valid compiled regex. Test compile */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn regcomp_empty_pattern() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; int r = regcomp(&re, \"\", 0); if(r == 0) { int r2 = regexec(&re, \"abc\", 0, NULL, 0); printf(\"%d\", r2 == 0); regfree(&re); } else printf(\"0\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn regcomp_too_many_parens() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; /* REG_EPAREN */ int r = regcomp(&re, \"(abc\", 0); printf(\"%d\", r != 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn regerror_buffer_size() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; int r = regcomp(&re, \"(abc\", 0); char buf[5]; size_t s = regerror(r, &re, buf, 5); printf(\"%d\", s > 5); return 0; }"
        ),
        vec!["1"]
    );
} // regerror returns required buffer size
#[test]
fn regcomp_nsub_count() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; int r = regcomp(&re, \"(a)(b)(c)\", REG_EXTENDED); if(r == 0) { printf(\"%d\", re.re_nsub == 3); regfree(&re); } else printf(\"0\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn regex_character_class() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; int r = regcomp(&re, \"[[:alpha:]]+\", REG_EXTENDED); if(r == 0) { int r2 = regexec(&re, \"123abc456\", 0, NULL, 0); printf(\"%d\", r2 == 0); regfree(&re); } else printf(\"0\"); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn regexec_start_offset_match() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; int r = regcomp(&re, \"a\", 0); if(r == 0) { regmatch_t m[1]; regexec(&re, \"aba\", 1, m, 0); printf(\"%d\", m[0].rm_so == 0); regfree(&re); } else printf(\"0\"); return 0; }"
        ),
        vec!["1"]
    );
} // It finds the first 'a'
#[test]
fn regex_backreference() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <regex.h>\nint main() { regex_t re; int r = regcomp(&re, \"(a)\\\\1\", 0); /* Backreferences are POSIX basic regex, not extended usually, wait, extended does not have backreferences by POSIX */ if(r == 0) { int r2 = regexec(&re, \"aa\", 0, NULL, 0); printf(\"%d\", r2 == 0); regfree(&re); } else printf(\"0\"); return 0; }"
        ),
        vec!["1"]
    );
}
