use super::helpers::*;

#[test]
fn fputc_writes_to_stdout() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    fputc('A', stdout);
    fputc('\n', stdout);
    return 0;
}
"#,
        &["A"],
    );
}

#[test]
fn putchar_sequence() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    putchar('h');
    putchar('i');
    putchar('\n');
    return 0;
}
"#,
        &["hi"],
    );
}

#[test]
fn fgetc_from_string_file() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    FILE *f = fopen("/tmp/vybe_fgetc.txt", "w");
    fputs("abc", f);
    fclose(f);
    f = fopen("/tmp/vybe_fgetc.txt", "r");
    int c;
    while ((c = fgetc(f)) != EOF) putchar(c);
    putchar('\n');
    fclose(f);
    return 0;
}
"#,
        &["abc"],
    );
}

#[test]
fn ungetc_pushes_back() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    FILE *f = fopen("/tmp/vybe_ungetc.txt", "w");
    fputs("bc", f);
    fclose(f);
    f = fopen("/tmp/vybe_ungetc.txt", "r");
    int c = fgetc(f);
    ungetc('a', f);
    int first = fgetc(f);
    fclose(f);
    printf("%c%c\n", first, c);
    return 0;
}
"#,
        &["ab"],
    );
}

#[test]
fn fflush_stdout_no_error() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    printf("hello ");
    fflush(stdout);
    printf("world\n");
    return 0;
}
"#,
        &["hello world"],
    );
}

#[test]
fn getc_same_as_fgetc() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    FILE *f = fopen("/tmp/vybe_getc.txt", "w");
    fputs("xyz", f);
    fclose(f);
    f = fopen("/tmp/vybe_getc.txt", "r");
    printf("%c\n", getc(f));
    fclose(f);
    return 0;
}
"#,
        &["x"],
    );
}
