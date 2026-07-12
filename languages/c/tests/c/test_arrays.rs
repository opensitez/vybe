use super::helpers::*;

#[test]
fn array_init_and_access() {
    let out = run_prints(
        r#"
#include <stdio.h>
int main() {
    int arr[5] = {10, 20, 30, 40, 50};
    printf("%d\n", arr[0]);
    printf("%d\n", arr[4]);
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["10", "50"]);
}

#[test]
fn array_loop() {
    let out = run_prints(
        r#"
#include <stdio.h>
int main() {
    int arr[3] = {1, 2, 3};
    for (int i = 0; i < 3; i++) {
        printf("%d\n", arr[i]);
    }
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn array_sum() {
    let out = run_prints(
        r#"
#include <stdio.h>
int main() {
    int arr[5] = {1, 2, 3, 4, 5};
    int sum = 0;
    for (int i = 0; i < 5; i++) {
        sum += arr[i];
    }
    printf("%d\n", sum);
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn multidim_array() {
    let out = run_prints(
        r#"
#include <stdio.h>
int main() {
    int m[2][3] = {{1,2,3},{4,5,6}};
    printf("%d\n", m[0][1]);
    printf("%d\n", m[1][2]);
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["2", "6"]);
}
