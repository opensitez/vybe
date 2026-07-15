use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn posix_dprintf_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <stdio.h>\n#include <unistd.h>\n#include <fcntl.h>\nint main() { int fd = open(\"test_dprintf.txt\", O_CREAT|O_WRONLY, 0644); if (fd != -1) { dprintf(fd, \"hello %d\", 123); close(fd); FILE *f = fopen(\"test_dprintf.txt\", \"r\"); char buf[20]; fgets(buf, sizeof(buf), f); printf(\"%s\", buf); fclose(f); } return 0; }"
        ),
        vec!["hello 123"]
    );
}
#[test]
fn posix_flockfile_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 199506L\n#include <stdio.h>\nint main() { FILE *f = fopen(\"test_flock.txt\", \"w+\"); if (f) { flockfile(f); fputs(\"locked\", f); funlockfile(f); rewind(f); char buf[10]; fgets(buf, sizeof(buf), f); printf(\"%s\", buf); fclose(f); } return 0; }"
        ),
        vec!["locked"]
    );
}
#[test]
fn posix_flockfile_recursive() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 199506L\n#include <stdio.h>\nint main() { FILE *f = fopen(\"test_flock2.txt\", \"w+\"); if (f) { flockfile(f); flockfile(f); fputs(\"ok\", f); funlockfile(f); funlockfile(f); printf(\"ok\"); fclose(f); } return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn posix_ftrylockfile_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 199506L\n#include <stdio.h>\nint main() { FILE *f = fopen(\"test_flock3.txt\", \"w+\"); if (f) { int res = ftrylockfile(f); printf(\"%d\", res == 0); if (res == 0) funlockfile(f); fclose(f); } return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn posix_getc_unlocked() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 199506L\n#include <stdio.h>\nint main() { FILE *f = fopen(\"test_getc_unlocked.txt\", \"w+\"); if (f) { fputs(\"A\", f); rewind(f); flockfile(f); printf(\"%c\", getc_unlocked(f)); funlockfile(f); fclose(f); } return 0; }"
        ),
        vec!["A"]
    );
}
#[test]
fn posix_putc_unlocked() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 199506L\n#include <stdio.h>\nint main() { FILE *f = fopen(\"test_putc_unlocked.txt\", \"w+\"); if (f) { flockfile(f); putc_unlocked('B', f); funlockfile(f); rewind(f); printf(\"%c\", fgetc(f)); fclose(f); } return 0; }"
        ),
        vec!["B"]
    );
}
