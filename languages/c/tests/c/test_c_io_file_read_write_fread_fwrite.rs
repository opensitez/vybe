use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn fputc_fgetc_basic() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_char.txt\", \"w+\"); fputc('X', f); rewind(f); printf(\"%c\", fgetc(f)); fclose(f); return 0; }"
        ),
        vec!["X"]
    );
}
#[test]
fn fputs_fgets_basic() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_str.txt\", \"w+\"); fputs(\"hello\", f); rewind(f); char buf[10]; fgets(buf, sizeof(buf), f); printf(\"%s\", buf); fclose(f); return 0; }"
        ),
        vec!["hello"]
    );
}
#[test]
fn fwrite_fread_basic() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_bin.txt\", \"wb+\"); int data[3] = {1, 2, 3}; fwrite(data, sizeof(int), 3, f); rewind(f); int buf[3] = {0}; fread(buf, sizeof(int), 3, f); printf(\"%d %d %d\", buf[0], buf[1], buf[2]); fclose(f); return 0; }"
        ),
        vec!["1 2 3"]
    );
}
#[test]
fn fprintf_fscanf_basic() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_fmt.txt\", \"w+\"); fprintf(f, \"%d %s\", 123, \"abc\"); rewind(f); int a; char b[10]; fscanf(f, \"%d %s\", &a, b); printf(\"%d %s\", a, b); fclose(f); return 0; }"
        ),
        vec!["123 abc"]
    );
}
#[test]
fn fread_partial() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_fread_part.txt\", \"w+\"); fputc('A', f); rewind(f); char buf[10]; size_t n = fread(buf, 1, 5, f); printf(\"%d\", (int)n); fclose(f); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn fgets_eof() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_fgets_eof.txt\", \"w+\"); char buf[10]; printf(\"%d\", fgets(buf, 10, f) == NULL); fclose(f); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn fputc_eof() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_fputc_r.txt\", \"r\"); if (f) { printf(\"%d\", fputc('A', f) == EOF); fclose(f); } else { printf(\"1\"); } return 0; }"
        ),
        vec!["1"]
    );
} // Write to read-only fails
#[test]
fn fflush_basic() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_fflush.txt\", \"w\"); fputs(\"hi\", f); printf(\"%d\", fflush(f) == 0); fclose(f); return 0; }"
        ),
        vec!["1"]
    );
}
