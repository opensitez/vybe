use super::helpers::*;

// scanf reads from stdin; these tests are limited to compile-only for stdin-based
// but we can test sscanf variants thoroughly since they read from strings
#[test]
fn sscanf_char_format() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    char c;
    sscanf("A", "%c", &c);
    printf("%c\n", c);
    return 0;
}
"#,
        &["A"],
    );
}

#[test]
fn sscanf_hex_format() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    int n;
    sscanf("0xff", "%i", &n);
    printf("%d\n", n);
    return 0;
}
"#,
        &["255"],
    );
}

#[test]
fn sscanf_octal_format() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    int n;
    sscanf("010", "%i", &n);
    printf("%d\n", n);
    return 0;
}
"#,
        &["8"],
    );
}

#[test]
fn sscanf_width_limit() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    char buf[4];
    sscanf("hello", "%3s", buf);
    printf("%s\n", buf);
    return 0;
}
"#,
        &["hel"],
    );
}

#[test]
fn sscanf_negative_number() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    int n;
    sscanf("-42", "%d", &n);
    printf("%d\n", n);
    return 0;
}
"#,
        &["-42"],
    );
}

#[test]
fn sscanf_unsigned_long() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    unsigned long n;
    sscanf("4000000000", "%lu", &n);
    printf("%lu\n", n);
    return 0;
}
"#,
        &["4000000000"],
    );
}
