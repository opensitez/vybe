use super::helpers::*;

#[test]
fn global_int_initialized_to_value() {
    assert_outputs(
        r#"
#include <stdio.h>
int global = 42;
int main() {
    printf("%d\n", global);
    return 0;
}
"#,
        &["42"],
    );
}

#[test]
fn global_modified_by_function() {
    assert_outputs(
        r#"
#include <stdio.h>
int counter = 0;
void increment() { counter++; }
int main() {
    increment();
    increment();
    increment();
    printf("%d\n", counter);
    return 0;
}
"#,
        &["3"],
    );
}

#[test]
fn global_array_initialized() {
    assert_outputs(
        r#"
#include <stdio.h>
int primes[] = {2, 3, 5, 7, 11};
int main() {
    printf("%d %d %d\n", primes[0], primes[2], primes[4]);
    return 0;
}
"#,
        &["2 5 11"],
    );
}

#[test]
fn global_string_constant() {
    assert_outputs(
        r#"
#include <stdio.h>
const char *greeting = "hello";
int main() {
    printf("%s\n", greeting);
    return 0;
}
"#,
        &["hello"],
    );
}

#[test]
fn global_struct_initialized() {
    assert_outputs(
        r#"
#include <stdio.h>
struct Config { int width; int height; };
struct Config screen = {1920, 1080};
int main() {
    printf("%d %d\n", screen.width, screen.height);
    return 0;
}
"#,
        &["1920 1080"],
    );
}

#[test]
fn global_zero_initialized_by_default() {
    assert_outputs(
        r#"
#include <stdio.h>
int uninitialized;
int main() {
    printf("%d\n", uninitialized);
    return 0;
}
"#,
        &["0"],
    );
}

#[test]
fn multiple_globals_independent() {
    assert_outputs(
        r#"
#include <stdio.h>
int x = 10;
int y = 20;
int z = 30;
int main() {
    printf("%d %d %d\n", x, y, z);
    return 0;
}
"#,
        &["10 20 30"],
    );
}
