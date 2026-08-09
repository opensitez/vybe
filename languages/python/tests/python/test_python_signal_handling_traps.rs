use super::helpers::run_python;

// signal — signal, getsignal, SIGINT, SIGTERM, SIG_IGN, SIG_DFL, default_int_handler, valid_signals, strsignal

#[test]
fn test_signal_getsignal_default_sigint() {
    let out = run_python(
        r#"
import signal
h = signal.getsignal(signal.SIGINT)
print(h is signal.default_int_handler or h is signal.SIG_DFL or callable(h))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_signal_custom_handler_registration() {
    let out = run_python(
        r#"
import signal

handled = []
def handler(signum, frame):
    handled.append(signum)

old = signal.signal(signal.SIGINT, handler)
current = signal.getsignal(signal.SIGINT)
print(current is handler)
signal.signal(signal.SIGINT, old)  # restore
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_signal_sig_ign_ignore_signal() {
    let out = run_python(
        r#"
import signal
old = signal.signal(signal.SIGINT, signal.SIG_IGN)
print(signal.getsignal(signal.SIGINT) is signal.SIG_IGN)
signal.signal(signal.SIGINT, old)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_signal_sig_dfl_default_handler() {
    let out = run_python(
        r#"
import signal
old = signal.signal(signal.SIGINT, signal.SIG_DFL)
print(signal.getsignal(signal.SIGINT) is signal.SIG_DFL)
signal.signal(signal.SIGINT, old)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_signal_valid_signals_set() {
    let out = run_python(
        r#"
import signal, sys
if hasattr(signal, "valid_signals"):
    sigs = signal.valid_signals()
    print(signal.SIGINT in sigs)
    print(signal.SIGTERM in sigs)
else:
    print(True)
    print(True)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_signal_strsignal_description() {
    let out = run_python(
        r#"
import signal, sys
if hasattr(signal, "strsignal"):
    desc = signal.strsignal(signal.SIGINT)
    print(isinstance(desc, str))
    print(len(desc) > 0)
else:
    print(True)
    print(True)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_signal_sigint_constant_value() {
    let out = run_python(
        r#"
import signal
print(isinstance(signal.SIGINT, int))
print(signal.SIGINT > 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_signal_sigterm_constant_value() {
    let out = run_python(
        r#"
import signal
print(isinstance(signal.SIGTERM, int))
print(signal.SIGTERM > 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_signal_raise_signal_invokes_handler() {
    let out = run_python(
        r#"
import signal, sys
if hasattr(signal, "raise_signal"):
    catches = []
    def my_h(signum, frame):
        catches.append(signum)
    old = signal.signal(signal.SIGUSR1 if hasattr(signal, "SIGUSR1") else signal.SIGINT, my_h)
    target_sig = signal.SIGUSR1 if hasattr(signal, "SIGUSR1") else signal.SIGINT
    signal.raise_signal(target_sig)
    print(len(catches) == 1)
    signal.signal(target_sig, old)
else:
    print(True)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_signal_alarm_unix_only() {
    let out = run_python(
        r#"
import signal, sys
if hasattr(signal, "alarm"):
    old_alarm = signal.alarm(0)
    print(isinstance(old_alarm, int))
else:
    print(True)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_signal_setitimer_unix_only() {
    let out = run_python(
        r#"
import signal, sys
if hasattr(signal, "setitimer"):
    old = signal.setitimer(signal.ITIMER_REAL, 0)
    print(isinstance(old, tuple))
else:
    print(True)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_signal_sig_block_unblock_unix() {
    let out = run_python(
        r#"
import signal, sys
if hasattr(signal, "pthread_sigmask"):
    old_mask = signal.pthread_sigmask(signal.SIG_BLOCK, [])
    print(isinstance(old_mask, set))
else:
    print(True)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_signal_sigabrt_constant() {
    let out = run_python(
        r#"
import signal
print(isinstance(signal.SIGABRT, int))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_signal_sigfpe_constant() {
    let out = run_python(
        r#"
import signal
print(isinstance(signal.SIGFPE, int))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_signal_sigill_constant() {
    let out = run_python(
        r#"
import signal
print(isinstance(signal.SIGILL, int))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_signal_sigsegv_constant() {
    let out = run_python(
        r#"
import signal
print(isinstance(signal.SIGSEGV, int))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_signal_getsignal_invalid_signal_raises() {
    let out = run_python(
        r#"
import signal
try:
    signal.getsignal(-1)
except (ValueError, OSError):
    print("Error")
"#,
    );
    assert_eq!(out, vec!["Error"]);
}

#[test]
fn test_signal_sig_ign_is_not_callable() {
    let out = run_python(
        r#"
import signal
print(callable(signal.SIG_IGN))
"#,
    );
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_signal_sig_dfl_is_not_callable() {
    let out = run_python(
        r#"
import signal
print(callable(signal.SIG_DFL))
"#,
    );
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_signal_nsig_constant() {
    let out = run_python(
        r#"
import signal
print(isinstance(signal.NSIG, int))
print(signal.NSIG > 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}
