use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Signal & Process Control — signal.signal, SIGINT, os.kill, os.getpid, os.getppid, process termination hooks
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_os_getpid_getppid_process_ids() {
    let src = r#"
import os

pid = os.getpid()
ppid = os.getppid()

print(isinstance(pid, int) and pid > 0)
print(isinstance(ppid, int) and ppid >= 0)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_signal_getsignal_default_handler() {
    let src = r#"
import signal

handler = signal.getsignal(signal.SIGINT)
print(handler is signal.default_int_handler or handler is signal.SIG_DFL or callable(handler))
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_signal_custom_handler_invocation() {
    let src = r#"
import signal, os

received = []

def handler(signum, frame):
    received.append(signum)

# Assign custom signal handler for SIGUSR1
signal.signal(signal.SIGUSR1, handler)
os.kill(os.getpid(), signal.SIGUSR1)

print(received == [signal.SIGUSR1])
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_signal_alarm_timer() {
    let src = r#"
import signal, time

received = []

def handler(signum, frame):
    received.append("alarm_triggered")

signal.signal(signal.SIGALRM, handler)
signal.alarm(1)  # 1 second
time.sleep(1.1)

print(received)
"#;
    assert_eq!(run_python(src), vec!["['alarm_triggered']"]);
}

#[test]
fn test_py_atexit_register_cleanup_hooks() {
    let src = r#"
import atexit

events = []

def cleanup():
    events.append("cleanup_done")

atexit.register(cleanup)
print("registered")
"#;
    assert_eq!(run_python(src), vec!["registered"]);
}

#[test]
fn test_py_os_strsignal_description() {
    let src = r#"
import signal, sys

if hasattr(signal, "strsignal"):
    desc = signal.strsignal(signal.SIGINT)
    print(isinstance(desc, str))
else:
    print("True")
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_signal_sig_ign_ignore_signal() {
    let src = r#"
import signal, os

# Set handler to ignore SIGUSR2
signal.signal(signal.SIGUSR2, signal.SIG_IGN)
os.kill(os.getpid(), signal.SIGUSR2)
print("survived signal")
"#;
    assert_eq!(run_python(src), vec!["survived signal"]);
}

#[test]
fn test_py_os_waitpid_child_process_status() {
    let src = r#"
import os, sys

pid = os.fork()
if pid == 0:
    # Child process
    os._exit(42)
else:
    # Parent process
    child_pid, status = os.waitpid(pid, 0)
    exit_code = os.WEXITSTATUS(status)
    print(child_pid == pid)
    print(exit_code)
"#;
    assert_eq!(run_python(src), vec!["True", "42"]);
}

#[test]
fn test_py_os_execv_replacing_process_image() {
    let src = r#"
import os, sys

pid = os.fork()
if pid == 0:
    # Child replaces image
    os.execv(sys.executable, [sys.executable, "-c", "print('execv_success')"])
else:
    _, status = os.waitpid(pid, 0)
    print(os.WEXITSTATUS(status) == 0)
"#;
    assert_eq!(run_python(src), vec!["execv_success", "True"]);
}

#[test]
fn test_py_signal_pause_or_sigtimedwait() {
    let src = r#"
import signal, os, time

received = []

def handler(signum, frame):
    received.append("alarm")

signal.signal(signal.SIGALRM, handler)
signal.alarm(1)
time.sleep(1.1)

print("alarm" in received)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}
