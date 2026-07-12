use super::helpers::*;

#[test]
fn fopen_fclose_basic() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    FILE *f = fopen("/tmp/vybe_test_fopen.txt", "w");
    printf("%d\n", f != NULL ? 1 : 0);
    fclose(f);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn fputs_fgets_roundtrip() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    FILE *f = fopen("/tmp/vybe_test_fputs.txt", "w");
    fputs("hello world\n", f);
    fclose(f);
    f = fopen("/tmp/vybe_test_fputs.txt", "r");
    char buf[50];
    fgets(buf, sizeof(buf), f);
    fclose(f);
    printf("%s", buf);
    return 0;
}
"#,
        &["hello world"],
    );
}

#[test]
fn fprintf_writes_formatted() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    FILE *f = fopen("/tmp/vybe_test_fprintf.txt", "w");
    fprintf(f, "%d %s\n", 42, "test");
    fclose(f);
    f = fopen("/tmp/vybe_test_fprintf.txt", "r");
    char buf[50];
    fgets(buf, sizeof(buf), f);
    fclose(f);
    printf("%s", buf);
    return 0;
}
"#,
        &["42 test"],
    );
}

#[test]
fn fwrite_fread_binary() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    int data[3] = {10, 20, 30};
    FILE *f = fopen("/tmp/vybe_test_fwrite.bin", "wb");
    fwrite(data, sizeof(int), 3, f);
    fclose(f);
    int readback[3] = {0, 0, 0};
    f = fopen("/tmp/vybe_test_fwrite.bin", "rb");
    fread(readback, sizeof(int), 3, f);
    fclose(f);
    printf("%d %d %d\n", readback[0], readback[1], readback[2]);
    return 0;
}
"#,
        &["10 20 30"],
    );
}

#[test]
fn fseek_ftell_position() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    FILE *f = fopen("/tmp/vybe_test_fseek.txt", "w");
    fputs("abcdef", f);
    fclose(f);
    f = fopen("/tmp/vybe_test_fseek.txt", "r");
    fseek(f, 3, SEEK_SET);
    long pos = ftell(f);
    printf("%ld\n", pos);
    fclose(f);
    return 0;
}
"#,
        &["3"],
    );
}

#[test]
fn feof_detects_end_of_file() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    FILE *f = fopen("/tmp/vybe_test_feof.txt", "w");
    fputs("x", f);
    fclose(f);
    f = fopen("/tmp/vybe_test_feof.txt", "r");
    fgetc(f);
    fgetc(f);
    printf("%d\n", feof(f) ? 1 : 0);
    fclose(f);
    return 0;
}
"#,
        &["1"],
    );
}

#[test]
fn rewind_resets_file_position() {
    assert_outputs(
        r#"
#include <stdio.h>
int main() {
    FILE *f = fopen("/tmp/vybe_test_rewind.txt", "w");
    fputs("hello", f);
    fclose(f);
    f = fopen("/tmp/vybe_test_rewind.txt", "r");
    char buf1[10], buf2[10];
    fgets(buf1, sizeof(buf1), f);
    rewind(f);
    fgets(buf2, sizeof(buf2), f);
    fclose(f);
    printf("%d\n", strcmp(buf1, buf2) == 0 ? 1 : 0);
    return 0;
}
"#,
        &["1"],
    );
}
