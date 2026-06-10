use super::helpers::*;

// C11 _Atomic types and stdatomic.h operations
#[test]
fn atomic_int_load_store() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdatomic.h>
int main() {
    _Atomic int x = 0;
    atomic_store(&x, 42);
    printf("%d\n", atomic_load(&x));
    return 0;
}
"#,
        &["42"],
    );
}

#[test]
fn atomic_fetch_add() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdatomic.h>
int main() {
    atomic_int x = 10;
    int old = atomic_fetch_add(&x, 5);
    printf("%d %d\n", old, atomic_load(&x));
    return 0;
}
"#,
        &["10 15"],
    );
}

#[test]
fn atomic_fetch_sub() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdatomic.h>
int main() {
    atomic_int x = 100;
    int old = atomic_fetch_sub(&x, 30);
    printf("%d %d\n", old, atomic_load(&x));
    return 0;
}
"#,
        &["100 70"],
    );
}

#[test]
fn atomic_compare_exchange_success() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdatomic.h>
int main() {
    atomic_int x = 5;
    int expected = 5;
    int ok = atomic_compare_exchange_strong(&x, &expected, 10);
    printf("%d %d\n", ok, atomic_load(&x));
    return 0;
}
"#,
        &["1 10"],
    );
}

#[test]
fn atomic_compare_exchange_failure() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdatomic.h>
int main() {
    atomic_int x = 5;
    int expected = 99;
    int ok = atomic_compare_exchange_strong(&x, &expected, 10);
    printf("%d %d\n", ok, atomic_load(&x));
    return 0;
}
"#,
        &["0 5"],
    );
}

#[test]
fn atomic_flag_test_and_set() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdatomic.h>
int main() {
    atomic_flag f = ATOMIC_FLAG_INIT;
    int first = atomic_flag_test_and_set(&f);
    int second = atomic_flag_test_and_set(&f);
    printf("%d %d\n", first, second);
    return 0;
}
"#,
        &["0 1"],
    );
}

#[test]
fn atomic_int_type_qualifier() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    _Atomic int x = 5;
    x++;
    printf("%d\n", x);
    return 0;
}
"#,
        &["6"],
    );
}
