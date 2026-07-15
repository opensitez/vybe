use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn ferror_basic() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_ferror.txt\", \"r\"); if (!f) { printf(\"skip\"); return 0; } fputc('X', f); printf(\"%d\", ferror(f) != 0); fclose(f); return 0; }"
        ),
        vec!["1"]
    );
} // writing to read-only sets error
#[test]
fn ferror_clearerr() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_clearerr.txt\", \"r\"); if (!f) return 0; fputc('X', f); int err1 = ferror(f) != 0; clearerr(f); int err2 = ferror(f) != 0; printf(\"%d %d\", err1, err2); fclose(f); return 0; }"
        ),
        vec!["1 0"]
    );
}
#[test]
fn feof_basic() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_feof.txt\", \"w+\"); fputs(\"A\", f); rewind(f); fgetc(f); int eof1 = feof(f); fgetc(f); int eof2 = feof(f); printf(\"%d %d\", eof1 != 0, eof2 != 0); fclose(f); return 0; }"
        ),
        vec!["0 1"]
    );
}
#[test]
fn feof_clearerr() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_feof2.txt\", \"w+\"); fgetc(f); int eof1 = feof(f); clearerr(f); int eof2 = feof(f); printf(\"%d %d\", eof1 != 0, eof2 != 0); fclose(f); return 0; }"
        ),
        vec!["1 0"]
    );
}
#[test]
fn perror_basic() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"/invalid_path_that_doesnt_exist_1234\", \"r\"); if (!f) { perror(\"MyError\"); } printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
} // perror goes to stderr, ignored by run_prints
#[test]
fn ferror_no_error() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_noerr.txt\", \"w+\"); printf(\"%d\", ferror(f)); fclose(f); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn feof_no_eof() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_noeof.txt\", \"w+\"); printf(\"%d\", feof(f)); fclose(f); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn clearerr_no_op() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_clearnoop.txt\", \"w+\"); clearerr(f); printf(\"%d %d\", ferror(f), feof(f)); fclose(f); return 0; }"
        ),
        vec!["0 0"]
    );
}
#[test]
fn ferror_stdin() {
    assert_eq!(
        run_c("int main() { printf(\"%d\", ferror(stdin)); return 0; }"),
        vec!["0"]
    );
}
#[test]
fn feof_stdin() {
    assert_eq!(
        run_c("int main() { printf(\"%d\", feof(stdin)); return 0; }"),
        vec!["0"]
    );
}
#[test]
fn clearerr_stdin() {
    assert_eq!(
        run_c("int main() { clearerr(stdin); printf(\"ok\"); return 0; }"),
        vec!["ok"]
    );
}
#[test]
fn perror_null() {
    assert_eq!(
        run_c("int main() { fopen(\"/invalid\", \"r\"); perror(NULL); printf(\"ok\"); return 0; }"),
        vec!["ok"]
    );
}
#[test]
fn perror_empty() {
    assert_eq!(
        run_c("int main() { fopen(\"/invalid\", \"r\"); perror(\"\"); printf(\"ok\"); return 0; }"),
        vec!["ok"]
    );
}
#[test]
fn ferror_after_fclose() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_fclose.txt\", \"w\"); fclose(f); /* ferror(f) is UB, don't test */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn feof_after_fseek() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_feof_seek.txt\", \"w+\"); fgetc(f); int eof1 = feof(f); fseek(f, 0, SEEK_SET); int eof2 = feof(f); printf(\"%d %d\", eof1 != 0, eof2 != 0); fclose(f); return 0; }"
        ),
        vec!["1 0"]
    );
} // fseek clears EOF
#[test]
fn feof_after_rewind() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_feof_rewind.txt\", \"w+\"); fgetc(f); int eof1 = feof(f); rewind(f); int eof2 = feof(f); printf(\"%d %d\", eof1 != 0, eof2 != 0); fclose(f); return 0; }"
        ),
        vec!["1 0"]
    );
} // rewind clears EOF
#[test]
fn ferror_after_fseek() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_ferr_seek.txt\", \"r\"); if (!f) return 0; fputc('X', f); int err1 = ferror(f) != 0; fseek(f, 0, SEEK_SET); int err2 = ferror(f) != 0; printf(\"%d %d\", err1, err2); fclose(f); return 0; }"
        ),
        vec!["1 1"]
    );
} // fseek does NOT clear error indicator
#[test]
fn ferror_after_rewind() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_ferr_rewind.txt\", \"r\"); if (!f) return 0; fputc('X', f); int err1 = ferror(f) != 0; rewind(f); int err2 = ferror(f) != 0; printf(\"%d %d\", err1, err2); fclose(f); return 0; }"
        ),
        vec!["1 0"]
    );
} // rewind DOES clear error indicator
#[test]
fn ferror_stdout() {
    assert_eq!(
        run_c("int main() { printf(\"%d\", ferror(stdout)); return 0; }"),
        vec!["0"]
    );
}
#[test]
fn clearerr_stdout() {
    assert_eq!(
        run_c("int main() { clearerr(stdout); printf(\"ok\"); return 0; }"),
        vec!["ok"]
    );
}
