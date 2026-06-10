use super::helpers::*;

#[test]
fn time_returns_positive_value() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <time.h>
int main() {
    time_t t = time(NULL);
    printf("%d\n", t > 0 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn clock_returns_nonnegative() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <time.h>
int main() {
    clock_t c = clock();
    printf("%d\n", c >= 0 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn difftime_positive_for_later() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <time.h>
int main() {
    time_t t1 = 1000;
    time_t t2 = 2000;
    double diff = difftime(t2, t1);
    printf("%.0f\n", diff);
    return 0;
}
"#,
        &["1000"],
    );
}

#[test]
fn gmtime_epoch_fields() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <time.h>
int main() {
    time_t epoch = 0;
    struct tm *t = gmtime(&epoch);
    printf("%d %d %d\n", t->tm_year + 1900, t->tm_mon + 1, t->tm_mday);
    return 0;
}
"#,
        &["1970 1 1"],
    );
}

#[test]
fn localtime_returns_struct() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <time.h>
int main() {
    time_t t = time(NULL);
    struct tm *lt = localtime(&t);
    printf("%d\n", lt->tm_year + 1900 >= 2024 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn mktime_roundtrip() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <time.h>
int main() {
    struct tm t = {0};
    t.tm_year = 70; t.tm_mon = 0; t.tm_mday = 1;
    time_t epoch = mktime(&t);
    printf("%d\n", epoch >= 0 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn strftime_formats_date() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <time.h>
#include <string.h>
int main() {
    time_t epoch = 0;
    struct tm *t = gmtime(&epoch);
    char buf[64];
    strftime(buf, sizeof(buf), "%Y-%m-%d", t);
    printf("%s\n", buf);
    return 0;
}
"#,
        &["1970-01-01"],
    );
}

#[test]
fn clocks_per_sec_positive() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <time.h>
int main() {
    printf("%d\n", CLOCKS_PER_SEC > 0 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}
