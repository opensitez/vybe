use super::helpers::*;

#[test]
fn math_sqrt() {
    let out = run_prints(
        r#"
#include <stdio.h>
#include <math.h>
int main() {
    printf("%.1f\n", sqrt(16.0));
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["4.0"]);
}

#[test]
fn math_abs() {
    let out = run_prints(
        r#"
#include <stdio.h>
#include <math.h>
int main() {
    printf("%.1f\n", fabs(-7.5));
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["7.5"]);
}

#[test]
fn math_pow() {
    let out = run_prints(
        r#"
#include <stdio.h>
#include <math.h>
int main() {
    printf("%.0f\n", pow(2.0, 10.0));
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["1024"]);
}

#[test]
fn math_floor_ceil() {
    let out = run_prints(
        r#"
#include <stdio.h>
#include <math.h>
int main() {
    printf("%.0f\n", floor(3.7));
    printf("%.0f\n", ceil(3.2));
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn stdlib_abs() {
    let out = run_prints(
        r#"
#include <stdio.h>
#include <stdlib.h>
int main() {
    printf("%d\n", abs(-42));
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["42"]);
}
