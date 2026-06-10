use super::helpers::*;

#[test]
fn vla_basic_declaration() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    int n = 5;
    int arr[n];
    for (int i = 0; i < n; i++) arr[i] = i * 2;
    printf("%d %d %d\n", arr[0], arr[2], arr[4]);
    return 0;
}
"#,
        &["0 4 8"],
    );
}

#[test]
fn vla_size_from_function_arg() {
    assert_outputs(
        r#"
#include <stdio.h>
void fill(int n) {
    int arr[n];
    for (int i = 0; i < n; i++) arr[i] = i + 1;
    for (int i = 0; i < n; i++) printf("%d\n", arr[i]);
}
int main() {
    fill(3);
    return 0;
}
"#,
        &["1", "2", "3"],
    );
}

#[test]
fn vla_size_at_runtime() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    int n;
    scanf("%d", &n);
    int arr[n];
    return 0;
}
"#,
        &[],
    );
}

#[test]
fn vla_different_sizes_same_scope() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    int n = 3;
    int m = 4;
    int a[n];
    int b[m];
    for (int i = 0; i < n; i++) a[i] = i;
    for (int i = 0; i < m; i++) b[i] = i * 10;
    printf("%d %d\n", a[2], b[3]);
    return 0;
}
"#,
        &["2 30"],
    );
}
