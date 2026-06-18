use super::helpers::*;

// POSIX <regex.h>: regcomp/regexec/regfree on the ECMA RegExp surface.

#[test]
fn regexec_reports_whole_match_offsets() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <regex.h>
int main() {
    regex_t re;
    regmatch_t m[1];
    regcomp(&re, "[0-9]+", REG_EXTENDED);
    regexec(&re, "ab123cd", 1, m, 0);
    printf("%d %d\n", m[0].rm_so, m[0].rm_eo);
    regfree(&re);
    return 0;
}
"#,
        &["2 5"],
    );
}

#[test]
fn regexec_reports_capture_group_offsets() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <regex.h>
int main() {
    regex_t re;
    regmatch_t m[3];
    regcomp(&re, "([0-9]+)-([0-9]+)", REG_EXTENDED);
    regexec(&re, "x 42-99 y", 3, m, 0);
    printf("%d-%d %d-%d %d-%d\n",
        m[0].rm_so, m[0].rm_eo, m[1].rm_so, m[1].rm_eo, m[2].rm_so, m[2].rm_eo);
    regfree(&re);
    return 0;
}
"#,
        &["2-7 2-4 5-7"],
    );
}

#[test]
fn regexec_returns_nomatch_when_no_match() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <regex.h>
int main() {
    regex_t re;
    regmatch_t m[1];
    regcomp(&re, "[0-9]+", REG_EXTENDED);
    int rc = regexec(&re, "no digits", 1, m, 0);
    printf("%d\n", rc == REG_NOMATCH ? 1 : 0);
    regfree(&re);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn regcomp_sets_re_nsub_to_group_count() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <regex.h>
int main() {
    regex_t re;
    regcomp(&re, "(a)(b)c", REG_EXTENDED);
    printf("%d\n", (int)re.re_nsub);
    regfree(&re);
    return 0;
}
"#,
        &["2"],
    );
}

#[test]
fn regcomp_rejects_invalid_pattern_with_message() {
    // The specific error code is implementation-defined; a non-zero return and a
    // non-empty regerror message are the portable guarantees.
    assert_outputs(
        r#"
#include <stdio.h>
#include <string.h>
#include <regex.h>
int main() {
    regex_t re;
    int rc = regcomp(&re, "a(b", REG_EXTENDED);
    char buf[64];
    regerror(rc, &re, buf, sizeof(buf));
    printf("%d %d\n", rc != 0 ? 1 : 0, strlen(buf) > 0 ? 1 : 0);
    return 0;
}
"#,
        &["1 1"],
    );
}

#[test]
fn regexec_nosub_still_reports_match() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <regex.h>
int main() {
    regex_t re;
    regmatch_t m[1];
    regcomp(&re, "foo", REG_EXTENDED | REG_NOSUB);
    int rc = regexec(&re, "a foo b", 1, m, 0);
    printf("%d\n", rc);
    regfree(&re);
    return 0;
}
"#,
        &["0"],
    );
}

#[test]
fn regcomp_icase_flag_matches_case_insensitively() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <regex.h>
int main() {
    regex_t re;
    regmatch_t m[1];
    regcomp(&re, "hello", REG_EXTENDED | REG_ICASE);
    int rc = regexec(&re, "say HELLO now", 1, m, 0);
    printf("%d %d\n", rc, m[0].rm_so);
    regfree(&re);
    return 0;
}
"#,
        &["0 4"],
    );
}
