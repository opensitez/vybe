use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn tmpfile_basic() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = tmpfile(); if (!f) return 1; fputs(\"hello\", f); rewind(f); char buf[10]; fgets(buf, sizeof(buf), f); printf(\"%s\", buf); fclose(f); return 0; }"
        ),
        vec!["hello"]
    );
}
#[test]
fn tmpnam_basic() {
    assert_eq!(
        run_c(
            "int main() { char buf[L_tmpnam]; char *p = tmpnam(buf); printf(\"%d\", p == buf); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn tmpnam_null() {
    assert_eq!(
        run_c("int main() { char *p = tmpnam(NULL); printf(\"%d\", p != NULL); return 0; }"),
        vec!["1"]
    );
} // Uses static buffer
#[test]
fn mkstemp_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\n#include <unistd.h>\nint main() { char tmpl[] = \"test_mkstemp_XXXXXX\"; int fd = mkstemp(tmpl); printf(\"%d %d\", fd >= 0, tmpl[14] != 'X'); if(fd>=0) { close(fd); unlink(tmpl); } return 0; }"
        ),
        vec!["1 1"]
    );
}
#[test]
fn tmpfile_multiple() {
    assert_eq!(
        run_c(
            "int main() { FILE *f1 = tmpfile(); FILE *f2 = tmpfile(); printf(\"%d\", f1 != NULL && f2 != NULL && f1 != f2); if(f1) fclose(f1); if(f2) fclose(f2); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn tmpfile_auto_remove() {
    assert_eq!(
        run_c(
            "int main() { /* We can't easily test deletion externally in one run_c, but we can assure it opens and writes */ FILE *f = tmpfile(); fputc('X', f); printf(\"ok\"); fclose(f); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn tmpnam_r_gnu() {
    assert_eq!(
        run_c(
            "#define _GNU_SOURCE\nint main() { char buf[L_tmpnam]; char *p = tmpnam_r(buf); printf(\"%d\", p == buf); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mkostemp_basic() {
    assert_eq!(
        run_c(
            "#define _GNU_SOURCE\n#include <stdlib.h>\n#include <unistd.h>\n#include <fcntl.h>\nint main() { char tmpl[] = \"test_mkost_XXXXXX\"; int fd = mkostemp(tmpl, O_APPEND); printf(\"%d\", fd >= 0); if(fd>=0) { close(fd); unlink(tmpl); } return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mkdtemp_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\n#include <unistd.h>\nint main() { char tmpl[] = \"test_mkdtemp_XXXXXX\"; char *p = mkdtemp(tmpl); printf(\"%d\", p == tmpl); if(p) rmdir(tmpl); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mktemp_basic() {
    assert_eq!(
        run_c(
            "#define _BSD_SOURCE\n#include <stdlib.h>\nint main() { char tmpl[] = \"test_mktemp_XXXXXX\"; char *p = mktemp(tmpl); printf(\"%d\", p == tmpl); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn mkstemp_file_contents() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\n#include <unistd.h>\nint main() { char tmpl[] = \"test_tmp_XXXXXX\"; int fd = mkstemp(tmpl); write(fd, \"data\", 4); close(fd); FILE *f = fopen(tmpl, \"r\"); char buf[10]; fgets(buf, sizeof(buf), f); printf(\"%s\", buf); fclose(f); unlink(tmpl); return 0; }"
        ),
        vec!["data"]
    );
}
#[test]
fn tmpfile_large_data() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = tmpfile(); for(int i=0; i<1000; i++) fputc('A', f); rewind(f); fseek(f, 0, SEEK_END); printf(\"%ld\", ftell(f)); fclose(f); return 0; }"
        ),
        vec!["1000"]
    );
}
#[test]
fn mkstemps_basic() {
    assert_eq!(
        run_c(
            "#define _BSD_SOURCE\n#include <stdlib.h>\n#include <unistd.h>\nint main() { char tmpl[] = \"test_mkstemps_XXXXXX.txt\"; int fd = mkstemps(tmpl, 4); printf(\"%d %c%c%c%c\", fd >= 0, tmpl[20], tmpl[21], tmpl[22], tmpl[23]); if(fd>=0) { close(fd); unlink(tmpl); } return 0; }"
        ),
        vec!["1 .txt"]
    );
}
#[test]
fn mkostemps_basic() {
    assert_eq!(
        run_c(
            "#define _GNU_SOURCE\n#include <stdlib.h>\n#include <unistd.h>\n#include <fcntl.h>\nint main() { char tmpl[] = \"test_mkost_XXXXXX.txt\"; int fd = mkostemps(tmpl, 4, O_SYNC); printf(\"%d\", fd >= 0); if(fd>=0){ close(fd); unlink(tmpl); } return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn tmpfile_seek_tell() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = tmpfile(); fputs(\"abc\", f); fseek(f, 1, SEEK_SET); printf(\"%c %ld\", fgetc(f), ftell(f)); fclose(f); return 0; }"
        ),
        vec!["b 2"]
    );
}
#[test]
fn tmpfile_binary() {
    assert_eq!(
        run_c(
            "int main() { FILE *f = tmpfile(); fputc(0, f); fputc(255, f); rewind(f); printf(\"%d %d\", fgetc(f), fgetc(f)); fclose(f); return 0; }"
        ),
        vec!["0 255"]
    );
} // tmpfile is opened as "wb+"
#[test]
fn tmpnam_repeated_calls() {
    assert_eq!(
        run_c(
            "int main() { char b1[L_tmpnam]; char b2[L_tmpnam]; tmpnam(b1); tmpnam(b2); /* might be same or different, but shouldn't crash */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn mkstemp_too_short() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\nint main() { char tmpl[] = \"XX\"; int fd = mkstemp(tmpl); printf(\"%d\", fd == -1); return 0; }"
        ),
        vec!["1"]
    );
} // Needs at least 6 Xs
#[test]
fn mkdtemp_too_short() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <stdlib.h>\nint main() { char tmpl[] = \"XX\"; char *p = mkdtemp(tmpl); printf(\"%d\", p == NULL); return 0; }"
        ),
        vec!["1"]
    );
} // Needs at least 6 Xs
#[test]
fn tmpfile_fileno() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 1\nint main() { FILE *f = tmpfile(); printf(\"%d\", fileno(f) >= 0); fclose(f); return 0; }"
        ),
        vec!["1"]
    );
}
