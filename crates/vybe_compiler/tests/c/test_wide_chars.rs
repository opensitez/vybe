use super::helpers::*;

// Wide character support via <wchar.h>
#[test]
fn wchar_t_basic_value() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <wchar.h>
int main() {
    wchar_t c = L'A';
    printf("%d\n", c == 65 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn wcslen_basic() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <wchar.h>
int main() {
    wchar_t s[] = L"hello";
    printf("%d\n", (int)wcslen(s));
    return 0;
}
"#,
        &["5"],
    );
}

#[test]
fn wcscpy_basic() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <wchar.h>
int main() {
    wchar_t dst[10];
    wcscpy(dst, L"abc");
    printf("%d\n", wcslen(dst) == 3 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn wcscmp_equal() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <wchar.h>
int main() {
    printf("%d\n", wcscmp(L"abc", L"abc") == 0 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn wprintf_basic() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <wchar.h>
int main() {
    wprintf(L"%d\n", 42);
    return 0;
}
"#,
        &["42"],
    );
}

#[test]
fn swprintf_basic() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <wchar.h>
int main() {
    wchar_t buf[32];
    swprintf(buf, 32, L"val=%d", 99);
    printf("%d\n", (int)wcslen(buf) > 0 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn wchar_string_literal_compiles() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <wchar.h>
int main() {
    const wchar_t *s = L"wide string";
    printf("%d\n", s != NULL ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}
