use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn fgetpos_fsetpos_basic() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_pos.txt\", \"w+\"); fputs(\"abcdef\", f); rewind(f); fgetc(f); fgetc(f); fpos_t pos; fgetpos(f, &pos); fgetc(f); fgetc(f); fsetpos(f, &pos); printf(\"%c\", fgetc(f)); fclose(f); return 0; }"
        ),
        vec!["c"]
    );
}
#[test]
fn ftell_fseek_basic() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_ftell.txt\", \"w+\"); fputs(\"0123456789\", f); fseek(f, 5, SEEK_SET); printf(\"%ld %c\", ftell(f), fgetc(f)); fclose(f); return 0; }"
        ),
        vec!["5 5"]
    );
}
#[test]
fn fseek_end() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_seekend.txt\", \"w+\"); fputs(\"hello\", f); fseek(f, 0, SEEK_END); printf(\"%ld\", ftell(f)); fclose(f); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn fseek_cur() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_seekcur.txt\", \"w+\"); fputs(\"012345\", f); fseek(f, 2, SEEK_SET); fseek(f, 2, SEEK_CUR); printf(\"%c\", fgetc(f)); fclose(f); return 0; }"
        ),
        vec!["4"]
    );
}
#[test]
fn rewind_basic() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_rewind.txt\", \"w+\"); fputs(\"abc\", f); rewind(f); printf(\"%c %ld\", fgetc(f), ftell(f)); fclose(f); return 0; }"
        ),
        vec!["a 1"]
    );
}
#[test]
fn ftello_fseeko_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\nint main() { FILE *f = fopen(\"test_ftello.txt\", \"w+\"); fputs(\"0123\", f); fseeko(f, 2, SEEK_SET); printf(\"%ld\", (long)ftello(f)); fclose(f); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn fsetpos_to_beginning() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_pos_beg.txt\", \"w+\"); fpos_t pos; fgetpos(f, &pos); fputs(\"hello\", f); fsetpos(f, &pos); printf(\"%c\", fgetc(f)); fclose(f); return 0; }"
        ),
        vec!["h"]
    );
}
#[test]
fn fgetpos_at_eof() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_pos_eof.txt\", \"w+\"); fputs(\"abc\", f); fpos_t pos; fgetpos(f, &pos); fsetpos(f, &pos); printf(\"%d\", feof(f)); fclose(f); return 0; }"
        ),
        vec!["0"]
    );
} // Doesn't inherently set EOF until read
#[test]
fn fseek_past_end_write() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_seek_past.txt\", \"w+\"); fseek(f, 3, SEEK_SET); fputc('X', f); rewind(f); printf(\"%d %d %d %c\", fgetc(f), fgetc(f), fgetc(f), fgetc(f)); fclose(f); return 0; }"
        ),
        vec!["0 0 0 X"]
    );
}
#[test]
fn fseek_negative_cur() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_seek_neg.txt\", \"w+\"); fputs(\"012345\", f); fseek(f, 5, SEEK_SET); fseek(f, -3, SEEK_CUR); printf(\"%c\", fgetc(f)); fclose(f); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn fseek_negative_end() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_seek_negend.txt\", \"w+\"); fputs(\"012345\", f); fseek(f, -2, SEEK_END); printf(\"%c\", fgetc(f)); fclose(f); return 0; }"
        ),
        vec!["4"]
    );
}
#[test]
fn fgetpos_invalid() {
    assert_eq!(
        run_c(
            "int main() { int res = fgetpos(NULL, NULL); /* usually crashes, but we check compiler handling. Let's test a closed file instead */ FILE *f = fopen(\"test_inv.txt\", \"w\"); fclose(f); fpos_t pos; printf(\"%d\", fgetpos(f, &pos) != 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn fsetpos_invalid() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_inv2.txt\", \"w\"); fclose(f); fpos_t pos; printf(\"%d\", fsetpos(f, &pos) != 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn ftell_append_mode() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_append_tell.txt\", \"w\"); fputs(\"123\", f); fclose(f); f = fopen(\"test_append_tell.txt\", \"a\"); printf(\"%ld\", ftell(f)); fclose(f); return 0; }"
        ),
        vec!["3"]
    );
} // On some systems this is 0 until write, but usually reflects end
#[test]
fn fseek_append_mode_write() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_append_seek.txt\", \"w\"); fputs(\"123\", f); fclose(f); f = fopen(\"test_append_seek.txt\", \"a+\"); fseek(f, 0, SEEK_SET); fputs(\"45\", f); rewind(f); char buf[10]; fgets(buf, sizeof(buf), f); printf(\"%s\", buf); fclose(f); return 0; }"
        ),
        vec!["12345"]
    );
} // Append forces writes to end, regardless of fseek
#[test]
fn rewind_clears_error() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_rewind_err.txt\", \"r\"); if (!f) return 0; fputc('X', f); int e1 = ferror(f) != 0; rewind(f); int e2 = ferror(f) != 0; printf(\"%d %d\", e1, e2); fclose(f); return 0; }"
        ),
        vec!["1 0"]
    );
}
#[test]
fn rewind_clears_eof() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_rewind_eof.txt\", \"w+\"); fgetc(f); int e1 = feof(f) != 0; rewind(f); int e2 = feof(f) != 0; printf(\"%d %d\", e1, e2); fclose(f); return 0; }"
        ),
        vec!["1 0"]
    );
}
#[test]
fn fseek_clears_eof() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_seek_eof.txt\", \"w+\"); fgetc(f); int e1 = feof(f) != 0; fseek(f, 0, SEEK_SET); int e2 = feof(f) != 0; printf(\"%d %d\", e1, e2); fclose(f); return 0; }"
        ),
        vec!["1 0"]
    );
}
#[test]
fn fseek_does_not_clear_error() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = fopen(\"test_seek_err.txt\", \"r\"); if (!f) return 0; fputc('X', f); int e1 = ferror(f) != 0; fseek(f, 0, SEEK_SET); int e2 = ferror(f) != 0; printf(\"%d %d\", e1, e2); fclose(f); return 0; }"
        ),
        vec!["1 1"]
    );
}
#[test]
fn ftell_stdout() {
    assert_eq!(
        run_c(
            "int main() { /* ftell on pipes/terminals returns -1 and sets errno */ long pos = ftell(stdout); printf(\"%d\", pos == -1 || pos >= 0); return 0; }"
        ),
        vec!["1"]
    );
}
