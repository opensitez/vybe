use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn signal_basic_handler() {
    assert_eq!(
        run_c(
            "#include <signal.h>\n#include <stdlib.h>\nvoid h(int sig) { printf(\"caught %d\", sig); exit(0); }\nint main() { signal(SIGINT, h); raise(SIGINT); return 0; }"
        ),
        vec!["caught 2"]
    );
} // SIGINT is usually 2
#[test]
fn signal_ignore() {
    assert_eq!(
        run_c(
            "#include <signal.h>\nint main() { signal(SIGUSR1, SIG_IGN); raise(SIGUSR1); printf(\"ignored\"); return 0; }"
        ),
        vec!["ignored"]
    );
}
#[test]
fn signal_default() {
    assert_eq!(
        run_c(
            "#include <signal.h>\nint main() { signal(SIGUSR1, SIG_DFL); /* raise(SIGUSR1) would terminate */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn raise_basic() {
    assert_eq!(
        run_c(
            "#include <signal.h>\nint main() { int r = raise(0); /* 0 is valid to test, does nothing but checks permissions */ printf(\"%d\", r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn signal_return_old_handler() {
    assert_eq!(
        run_c(
            "#include <signal.h>\nvoid h(int s) {}\nint main() { void (*old)(int) = signal(SIGUSR1, h); void (*old2)(int) = signal(SIGUSR1, SIG_IGN); printf(\"%d\", old2 == h); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn sig_atomic_t_usage() {
    assert_eq!(
        run_c(
            "#include <signal.h>\nvolatile sig_atomic_t flag = 0;\nvoid h(int s) { flag = 1; }\nint main() { signal(SIGUSR1, h); raise(SIGUSR1); printf(\"%d\", flag); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn kill_self() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <signal.h>\n#include <unistd.h>\nvoid h(int s) { printf(\"killed\"); _exit(0); }\nint main() { signal(SIGUSR1, h); kill(getpid(), SIGUSR1); return 0; }"
        ),
        vec!["killed"]
    );
}
#[test]
fn kill_invalid_pid() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <signal.h>\nint main() { int r = kill(-99999, 0); printf(\"%d\", r == -1); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn alarm_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\n#include <signal.h>\nvoid h(int s) { printf(\"alarm\"); _exit(0); }\nint main() { signal(SIGALRM, h); alarm(1); pause(); return 0; }"
        ),
        vec!["alarm"]
    );
}
#[test]
fn alarm_cancel() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { alarm(10); int rem = alarm(0); printf(\"%d\", rem > 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pause_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\n#include <signal.h>\n#include <sys/wait.h>\nvoid h(int s) { _exit(5); }\nint main() { pid_t p = fork(); if(p==0) { signal(SIGUSR1, h); pause(); _exit(0); } sleep(1); kill(p, SIGUSR1); int st; waitpid(p, &st, 0); printf(\"%d\", WEXITSTATUS(st)); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn signal_error_return() {
    assert_eq!(
        run_c(
            "#include <signal.h>\nint main() { void (*r)(int) = signal(99999, SIG_IGN); printf(\"%d\", r == SIG_ERR); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn raise_sigill_catch() {
    assert_eq!(
        run_c(
            "#include <signal.h>\n#include <stdlib.h>\nvoid h(int s) { printf(\"ill\"); exit(0); }\nint main() { signal(SIGILL, h); raise(SIGILL); return 0; }"
        ),
        vec!["ill"]
    );
}
#[test]
fn raise_sigfpe_catch() {
    assert_eq!(
        run_c(
            "#include <signal.h>\n#include <stdlib.h>\nvoid h(int s) { printf(\"fpe\"); exit(0); }\nint main() { signal(SIGFPE, h); raise(SIGFPE); return 0; }"
        ),
        vec!["fpe"]
    );
}
#[test]
fn raise_sigsegv_catch() {
    assert_eq!(
        run_c(
            "#include <signal.h>\n#include <stdlib.h>\nvoid h(int s) { printf(\"segv\"); exit(0); }\nint main() { signal(SIGSEGV, h); raise(SIGSEGV); return 0; }"
        ),
        vec!["segv"]
    );
}
#[test]
fn kill_group() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <signal.h>\n#include <unistd.h>\nint main() { int r = kill(0, 0); /* 0 pid means send to process group, 0 sig means check perms */ printf(\"%d\", r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn killpg_basic() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE 500\n#include <signal.h>\n#include <unistd.h>\nint main() { int r = killpg(getpgrp(), 0); printf(\"%d\", r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn sleep_interrupt() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\n#include <signal.h>\nvoid h(int s) {}\nint main() { signal(SIGALRM, h); alarm(1); int rem = sleep(10); printf(\"%d\", rem > 0); return 0; }"
        ),
        vec!["1"]
    );
} // sleep returns unslept seconds if interrupted
#[test]
fn sigqueue_compile() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <signal.h>\n#include <unistd.h>\nint main() { union sigval val; val.sival_int = 42; int r = sigqueue(getpid(), 0, val); printf(\"%d\", r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn signal_multiple() {
    assert_eq!(
        run_c(
            "#include <signal.h>\n#include <stdlib.h>\nvoid h1(int s) { printf(\"1\"); exit(0); }\nvoid h2(int s) { printf(\"2\"); }\nint main() { signal(SIGUSR1, h1); signal(SIGUSR2, h2); raise(SIGUSR2); raise(SIGUSR1); return 0; }"
        ),
        vec!["21"]
    );
}
