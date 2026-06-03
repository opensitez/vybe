use super::helpers::*;

#[test]
fn atoi_parses_leading_whitespace_and_sign() {
    let out = run_prints(r#"
#include <stdio.h>
#include <stdlib.h>
int main() {
    printf("%d\n", atoi("   -42xyz"));
    return 0;
}
"#);
    assert_eq!(out, vec!["-42"]);
}

#[test]
fn atol_parses_positive_integer_text() {
    let out = run_prints(r#"
#include <stdio.h>
#include <stdlib.h>
int main() {
    printf("%d\n", atol("123456"));
    return 0;
}
"#);
    assert_eq!(out, vec!["123456"]);
}

#[test]
fn atof_parses_fractional_text() {
    let out = run_prints(r#"
#include <stdio.h>
#include <stdlib.h>
int main() {
    printf("%.2f\n", atof("98.125"));
    return 0;
}
"#);
    assert_eq!(out, vec!["98.12"]);
}

#[test]
fn labs_returns_absolute_value_for_negative_long() {
    let out = run_prints(r#"
#include <stdio.h>
#include <stdlib.h>
int main() {
    printf("%d\n", labs(-9001));
    return 0;
}
"#);
    assert_eq!(out, vec!["9001"]);
}