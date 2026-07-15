use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn setsid_basic() { assert_eq!(run_c("#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\n#include <sys/wait.h>\nint main() { pid_t p = fork(); if (p==0) { pid_t sid = setsid(); printf(\"%d\", sid > 0); _exit(0); } wait(NULL); return 0; }"), vec!["1"]); }
#[test] fn getsid_basic() { assert_eq!(run_c("#define _XOPEN_SOURCE 500\n#include <unistd.h>\n#include <sys/wait.h>\nint main() { pid_t p = fork(); if(p==0) { setsid(); printf(\"%d\", getsid(0) == getpid()); _exit(0); } wait(NULL); return 0; }"), vec!["1"]); }
#[test] fn setsid_already_group_leader() { assert_eq!(run_c("#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\n#include <sys/wait.h>\nint main() { pid_t p = fork(); if(p==0) { setsid(); int r = setsid(); printf(\"%d\", r == -1); _exit(0); } wait(NULL); return 0; }"), vec!["1"]); } // Cannot setsid if already process group leader
#[test] fn getpgrp_basic() { assert_eq!(run_c("#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { pid_t g = getpgrp(); printf(\"%d\", g > 0); return 0; }"), vec!["1"]); }
#[test] fn setpgid_basic() { assert_eq!(run_c("#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { int r = setpgid(0, 0); printf(\"%d\", r == 0); return 0; }"), vec!["1"]); }
#[test] fn getpgid_basic() { assert_eq!(run_c("#define _XOPEN_SOURCE 500\n#include <unistd.h>\nint main() { pid_t g = getpgid(0); printf(\"%d\", g == getpgrp()); return 0; }"), vec!["1"]); }
#[test] fn setpgrp_basic() { assert_eq!(run_c("#define _XOPEN_SOURCE 500\n#include <unistd.h>\nint main() { pid_t r = setpgrp(); printf(\"%d\", r == getpgrp()); return 0; }"), vec!["1"]); }
#[test] fn tcgetpgrp_compile() { assert_eq!(run_c("#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { pid_t g = tcgetpgrp(0); printf(\"%d\", g == -1 || g > 0); return 0; }"), vec!["1"]); }
#[test] fn tcsetpgrp_compile() { assert_eq!(run_c("#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { int r = tcsetpgrp(0, getpgrp()); printf(\"%d\", r == -1 || r == 0); return 0; }"), vec!["1"]); }
#[test] fn getsid_invalid_pid() { assert_eq!(run_c("#define _XOPEN_SOURCE 500\n#include <unistd.h>\nint main() { pid_t s = getsid(-99999); printf(\"%d\", s == -1); return 0; }"), vec!["1"]); }
#[test] fn setpgid_invalid_pid() { assert_eq!(run_c("#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { int r = setpgid(-99999, 0); printf(\"%d\", r == -1); return 0; }"), vec!["1"]); }
#[test] fn getpgid_invalid_pid() { assert_eq!(run_c("#define _XOPEN_SOURCE 500\n#include <unistd.h>\nint main() { pid_t g = getpgid(-99999); printf(\"%d\", g == -1); return 0; }"), vec!["1"]); }
#[test] fn chdir_root() { assert_eq!(run_c("#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { int r = chdir(\"/\"); printf(\"%d\", r == 0); return 0; }"), vec!["1"]); } // typical daemon step
#[test] fn umask_reset() { assert_eq!(run_c("#include <sys/stat.h>\nint main() { mode_t old = umask(0); umask(old); printf(\"ok\"); return 0; }"), vec!["ok"]); } // typical daemon step
#[test] fn daemon_func_gnu() { assert_eq!(run_c("#define _BSD_SOURCE\n#define _DEFAULT_SOURCE\n#include <unistd.h>\nint main() { /* don't actually daemonize the test runner */ printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn setpgid_child() { assert_eq!(run_c("#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\n#include <sys/wait.h>\nint main() { pid_t p = fork(); if(p==0) { setpgid(0,0); _exit(getpgrp() == getpid() ? 1 : 0); } int st; waitpid(p, &st, 0); printf(\"%d\", WEXITSTATUS(st)); return 0; }"), vec!["1"]); }
#[test] fn fork_twice_daemon_pattern() { assert_eq!(run_c("#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\n#include <sys/wait.h>\nint main() { pid_t p = fork(); if(p==0) { setsid(); pid_t p2 = fork(); if (p2==0) { _exit(0); } _exit(0); } wait(NULL); printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn getsid_parent() { assert_eq!(run_c("#define _XOPEN_SOURCE 500\n#include <unistd.h>\nint main() { pid_t s = getsid(getppid()); printf(\"%d\", s > 0); return 0; }"), vec!["1"]); }
#[test] fn tcgetsid_compile() { assert_eq!(run_c("#define _XOPEN_SOURCE 500\n#include <termios.h>\nint main() { pid_t s = tcgetsid(0); printf(\"%d\", s == -1 || s > 0); return 0; }"), vec!["1"]); }
#[test] fn getpgrp_no_args() { assert_eq!(run_c("#define _POSIX_C_SOURCE 200809L\n#include <unistd.h>\nint main() { /* getpgrp takes void in POSIX, int in BSD */ pid_t g = getpgrp(); printf(\"ok\"); return 0; }"), vec!["ok"]); }
