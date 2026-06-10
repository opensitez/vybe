use super::helpers::*;

// __LINE__ expands to current line number
#[test]
fn line_macro_increases() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    int a = __LINE__;
    int b = __LINE__;
    printf("%d\n", b > a ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

// __FILE__ expands to a non-empty string
#[test]
fn file_macro_is_string() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <string.h>
int main() {
    printf("%d\n", strlen(__FILE__) > 0 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

// __func__ inside a function
#[test]
fn func_macro_in_function() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <string.h>
void my_func() {
    printf("%s\n", __func__);
}
int main() {
    my_func();
    return 0;
}
"#,
        &["my_func"],
    );
}

// __DATE__ and __TIME__ are non-empty strings
#[test]
fn date_time_macros_are_strings() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <string.h>
int main() {
    printf("%d\n", strlen(__DATE__) > 0 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

// __LINE__ in a macro shows expansion site line
#[test]
fn line_in_macro_expansion() {
    assert_outputs(
        r#"
#include <stdio.h>
#define GET_LINE() __LINE__
int main() {
    int a = GET_LINE();
    int b = GET_LINE();
    printf("%d\n", b > a ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}
