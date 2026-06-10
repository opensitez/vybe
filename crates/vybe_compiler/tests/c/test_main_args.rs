use super::helpers::*;

#[test]
fn main_with_argc_argv_compiles() {
    assert_outputs(
        r#"
#include <stdio.h>
int main(int argc, char *argv[]) {
    printf("%d\n", argc >= 0 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn main_with_argc_and_argv_second_form() {
    assert_outputs(
        r#"
#include <stdio.h>
int main(int argc, char **argv) {
    printf("ok\n");
    return 0;
}
"#,
        &["ok"],
    );
}

#[test]
fn main_void_explicit() {
    assert_outputs(
        r#"
#include <stdio.h>
int main(void) {
    printf("void\n");
    return 0;
}
"#,
        &["void"],
    );
}

#[test]
fn main_return_zero_implicit_success() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    printf("done\n");
    return 0;
}
"#,
        &["done"],
    );
}
