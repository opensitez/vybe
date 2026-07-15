use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn kill_self_sigkill() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <signal.h>\n#include <unistd.h>\n#include <sys/wait.h>\nint main() { pid_t p = fork(); if (p==0) { kill(getpid(), SIGKILL); _exit(0); } int st; wait(&st); printf(\"%d\", WIFSIGNALED(st) && WTERMSIG(st) == SIGKILL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn kill_parent() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <signal.h>\n#include <unistd.h>\n#include <sys/wait.h>\nvoid h(int s) {}\nint main() { signal(SIGUSR1, h); pid_t p = fork(); if (p==0) { kill(getppid(), SIGUSR1); _exit(0); } wait(NULL); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn kill_negative_pid() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <signal.h>\n#include <unistd.h>\n#include <sys/wait.h>\nint main() { pid_t p = fork(); if (p==0) { setpgid(0,0); pid_t gp = getpgrp(); if(fork()==0) { kill(-gp, SIGKILL); _exit(0); } wait(NULL); _exit(0); } wait(NULL); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
} // SIGKILL to process group, both children die
#[test]
fn ualarm_basic() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE 500\n#include <unistd.h>\n#include <signal.h>\nvoid h(int s) { printf(\"ualarm\"); _exit(0); }\nint main() { signal(SIGALRM, h); ualarm(50000, 0); pause(); return 0; }"
        ),
        vec!["ualarm"]
    );
}
#[test]
fn ualarm_cancel() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE 500\n#include <unistd.h>\nint main() { ualarm(500000, 0); int r = ualarm(0, 0); printf(\"%d\", r > 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn setitimer_real() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/time.h>\n#include <signal.h>\n#include <unistd.h>\nvoid h(int s) { printf(\"timer\"); _exit(0); }\nint main() { signal(SIGALRM, h); struct itimerval it = {0}; it.it_value.tv_usec = 50000; setitimer(ITIMER_REAL, &it, NULL); pause(); return 0; }"
        ),
        vec!["timer"]
    );
}
#[test]
fn setitimer_virtual() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/time.h>\n#include <signal.h>\nvoid h(int s) { printf(\"vtimer\"); _exit(0); }\nint main() { signal(SIGVTALRM, h); struct itimerval it = {0}; it.it_value.tv_usec = 50000; setitimer(ITIMER_VIRTUAL, &it, NULL); while(1); return 0; }"
        ),
        vec!["vtimer"]
    );
} // Busy loop consumes virtual time
#[test]
fn getitimer_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/time.h>\nint main() { struct itimerval it = {0}; it.it_value.tv_sec = 10; setitimer(ITIMER_REAL, &it, NULL); struct itimerval old; getitimer(ITIMER_REAL, &old); printf(\"%d\", old.it_value.tv_sec > 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn siginterrupt_compile() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE 500\n#include <signal.h>\nint main() { int r = siginterrupt(SIGUSR1, 1); printf(\"%d\", r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn sigaltstack_compile() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <signal.h>\n#include <stdlib.h>\nint main() { stack_t ss; ss.ss_sp = malloc(SIGSTKSZ); ss.ss_size = SIGSTKSZ; ss.ss_flags = 0; int r = sigaltstack(&ss, NULL); printf(\"%d\", r == 0); free(ss.ss_sp); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn psignal_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <signal.h>\nint main() { /* psignal prints to stderr, just ensure compilation */ psignal(SIGINT, \"MySig\"); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn psiginfo_basic() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <signal.h>\nint main() { siginfo_t info = {0}; info.si_signo = SIGINT; psiginfo(&info, \"Info\"); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn kill_permission_denied() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <signal.h>\nint main() { /* init process pid=1 cannot be killed by unprivileged */ int r = kill(1, SIGKILL); printf(\"%d\", r == -1); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn alarm_overwrite() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { alarm(10); alarm(5); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn setitimer_prof() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/time.h>\n#include <signal.h>\nvoid h(int s) { printf(\"ptimer\"); _exit(0); }\nint main() { signal(SIGPROF, h); struct itimerval it = {0}; it.it_value.tv_usec = 50000; setitimer(ITIMER_PROF, &it, NULL); while(1); return 0; }"
        ),
        vec!["ptimer"]
    );
}
#[test]
fn getitimer_cleared() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <sys/time.h>\nint main() { struct itimerval it = {0}; setitimer(ITIMER_REAL, &it, NULL); getitimer(ITIMER_REAL, &it); printf(\"%d\", it.it_value.tv_sec == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn sigaltstack_disable() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <signal.h>\nint main() { stack_t ss; ss.ss_flags = SS_DISABLE; int r = sigaltstack(&ss, NULL); printf(\"%d\", r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn kill_zero_signal() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <signal.h>\n#include <unistd.h>\nint main() { int r = kill(getpid(), 0); printf(\"%d\", r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn sigaction_sa_onstack() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <signal.h>\nint main() { struct sigaction sa; sa.sa_handler = SIG_IGN; sigemptyset(&sa.sa_mask); sa.sa_flags = SA_ONSTACK; sigaction(SIGUSR1, &sa, NULL); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn sigaction_sa_resethand() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <signal.h>\nint main() { struct sigaction sa; sa.sa_handler = SIG_IGN; sigemptyset(&sa.sa_mask); sa.sa_flags = SA_RESETHAND; sigaction(SIGUSR1, &sa, NULL); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
