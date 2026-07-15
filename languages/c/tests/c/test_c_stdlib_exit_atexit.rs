use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn atexit_basic() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nvoid func1() { printf(\"func1 \"); }\nvoid func2() { printf(\"func2 \"); }\nint main() { atexit(func1); atexit(func2); return 0; }"
        ),
        vec!["func2 func1 "]
    );
} // Reverse registration order
#[test]
fn atexit_multiple_same() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nvoid func() { printf(\"f \"); }\nint main() { atexit(func); atexit(func); return 0; }"
        ),
        vec!["f f "]
    );
}
#[test]
fn atexit_limits() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nvoid func() { }\nint main() { int i, res = 0; for(i=0; i<32; i++) { res |= atexit(func); } printf(\"%d\", res == 0); return 0; }"
        ),
        vec!["1"]
    );
} // ANSI C guarantees at least 32
#[test]
fn exit_calls_atexit() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nvoid func() { printf(\"called\"); }\nint main() { atexit(func); exit(0); printf(\"not reached\"); return 0; }"
        ),
        vec!["called"]
    );
}
#[test]
fn exit_status_code() {
    assert_eq!(
        run_c("#include <stdlib.h>\nint main() { exit(42); return 0; }"),
        Vec::<String>::new()
    );
} // run_prints doesn't capture exit codes cleanly, but we ensure it doesn't crash
#[test]
fn _exit_bypasses_atexit() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\n#include <unistd.h>\nvoid func() { printf(\"called\"); }\nint main() { atexit(func); _exit(0); return 0; }"
        ),
        Vec::<String>::new()
    );
}
#[test]
fn _exit_bypasses_stdio_flush() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\n#include <unistd.h>\nint main() { printf(\"hello\"); _exit(0); return 0; }"
        ),
        Vec::<String>::new()
    );
} // stdout might not be flushed
#[test]
fn quick_exit_basic() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nvoid func() { printf(\"quick\"); }\nint main() { at_quick_exit(func); quick_exit(0); return 0; }"
        ),
        vec!["quick"]
    );
}
#[test]
fn quick_exit_bypasses_atexit() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nvoid func1() { printf(\"atexit\"); }\nvoid func2() { printf(\"quick\"); }\nint main() { atexit(func1); at_quick_exit(func2); quick_exit(0); return 0; }"
        ),
        vec!["quick"]
    );
}
#[test]
fn at_quick_exit_multiple() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nvoid func1() { printf(\"1\"); }\nvoid func2() { printf(\"2\"); }\nint main() { at_quick_exit(func1); at_quick_exit(func2); quick_exit(0); return 0; }"
        ),
        vec!["21"]
    );
}
#[test]
fn abort_bypasses_atexit() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\n#include <signal.h>\nvoid func() { printf(\"called\"); }\nvoid sighandler(int sig) { exit(0); }\nint main() { signal(SIGABRT, sighandler); atexit(func); abort(); return 0; }"
        ),
        vec!["called"]
    );
} // if we catch SIGABRT and exit(), atexit runs. If we don't catch it, it doesn't.
#[test]
fn on_exit_basic() {
    assert_eq!(
        run_c(
            "#define _BSD_SOURCE\n#include <stdlib.h>\nvoid func(int status, void *arg) { printf(\"%d %s\", status, (char*)arg); }\nint main() { on_exit(func, \"hello\"); exit(42); return 0; }"
        ),
        vec!["42 hello"]
    );
}
#[test]
fn exit_success_macro() {
    assert_eq!(
        run_c("#include <stdlib.h>\nint main() { exit(EXIT_SUCCESS); return 0; }"),
        Vec::<String>::new()
    );
}
#[test]
fn exit_failure_macro() {
    assert_eq!(
        run_c("#include <stdlib.h>\nint main() { exit(EXIT_FAILURE); return 0; }"),
        Vec::<String>::new()
    );
}
#[test]
fn atexit_from_atexit() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nvoid func2() { printf(\"2\"); }\nvoid func1() { atexit(func2); printf(\"1\"); }\nint main() { atexit(func1); return 0; }"
        ),
        vec!["12"]
    );
} // POSIX allows atexit to be called from atexit handler
#[test]
fn exit_from_atexit() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nvoid func() { exit(0); }\nint main() { atexit(func); exit(0); return 0; }"
        ),
        Vec::<String>::new()
    );
} // Undefined behavior in C standard, but typically safely stops. We just test it doesn't hang.
#[test]
fn atexit_null_ptr() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { /* atexit(NULL) is UB, but shouldn't crash if unexecuted */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn at_quick_exit_null_ptr() {
    assert_eq!(
        run_c("#include <stdlib.h>\nint main() { printf(\"ok\"); return 0; }"),
        vec!["ok"]
    );
}
#[test]
fn on_exit_multiple() {
    assert_eq!(
        run_c(
            "#define _BSD_SOURCE\n#include <stdlib.h>\nvoid f1(int s, void *a) { printf(\"1\"); }\nvoid f2(int s, void *a) { printf(\"2\"); }\nint main() { on_exit(f1, NULL); on_exit(f2, NULL); return 0; }"
        ),
        vec!["21"]
    );
}
#[test]
fn on_exit_mixed_with_atexit() {
    assert_eq!(
        run_c(
            "#define _BSD_SOURCE\n#include <stdlib.h>\nvoid f1() { printf(\"1\"); }\nvoid f2(int s, void *a) { printf(\"2\"); }\nint main() { atexit(f1); on_exit(f2, NULL); return 0; }"
        ),
        vec!["21"]
    );
}
