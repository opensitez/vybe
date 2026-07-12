use super::helpers::*;

#[test]
fn errno_set_on_domain_error() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <math.h>
#include <errno.h>
int main() {
    errno = 0;
    double result = sqrt(-1.0);
    printf("%d\n", errno != 0 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn strerror_returns_string() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <string.h>
#include <errno.h>
int main() {
    /* EINVAL = 22 on most platforms */
    char *msg = strerror(EINVAL);
    printf("%d\n", msg != NULL ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn errno_zero_after_success() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <math.h>
#include <errno.h>
int main() {
    errno = 0;
    double result = sqrt(4.0);
    printf("%.1f %d\n", result, errno);
    return 0;
}
"#,
        &["2.0 0"],
    );
}

#[test]
fn strerror_zero_is_success_string() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <string.h>
#include <errno.h>
int main() {
    char *msg = strerror(0);
    printf("%d\n", msg != NULL ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}
