use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Subprocess Process Management — subprocess.run, Popen, check_output, PIPE, DEVNULL, CalledProcessError, TimeoutExpired
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_subprocess_run_capture_stdout() {
    let src = r#"
import subprocess

result = subprocess.run(["python3", "-c", "print('hello from child')"], capture_output=True, text=True)
print(result.returncode)
print(result.stdout.strip())
"#;
    assert_eq!(run_python(src), vec!["0", "hello from child"]);
}

#[test]
fn test_py_subprocess_check_output_utf8_text() {
    let src = r#"
import subprocess

out = subprocess.check_output(["python3", "-c", "import sys; sys.stdout.write('output_string')"], text=True)
print(out)
"#;
    assert_eq!(run_python(src), vec!["output_string"]);
}

#[test]
fn test_py_subprocess_called_process_error_check() {
    let src = r#"
import subprocess

try:
    subprocess.run(["python3", "-c", "import sys; sys.exit(42)"], check=True)
except subprocess.CalledProcessError as e:
    print(f"CalledProcessError returncode={e.returncode}")
"#;
    assert_eq!(run_python(src), vec!["CalledProcessError returncode=42"]);
}

#[test]
fn test_py_subprocess_timeout_expired_handling() {
    let src = r#"
import subprocess

try:
    subprocess.run(["python3", "-c", "import time; time.sleep(10)"], timeout=0.01)
except subprocess.TimeoutExpired as e:
    print("TimeoutExpired caught")
"#;
    assert_eq!(run_python(src), vec!["TimeoutExpired caught"]);
}

#[test]
fn test_py_subprocess_popen_pipe_communication() {
    let src = r#"
import subprocess

proc = subprocess.Popen(
    ["python3", "-c", "import sys; data = sys.stdin.read(); sys.stdout.write(f'REPLY:{data}')"],
    stdin=subprocess.PIPE,
    stdout=subprocess.PIPE,
    text=True
)
stdout, _ = proc.communicate(input="REQUEST_DATA")
print(stdout)
print(proc.returncode)
"#;
    assert_eq!(run_python(src), vec!["REPLY:REQUEST_DATA", "0"]);
}

#[test]
fn test_py_subprocess_env_dictionary_pass() {
    let src = r#"
import subprocess, os

env = os.environ.copy()
env["CUSTOM_VAR"] = "custom_val_123"

result = subprocess.run(
    ["python3", "-c", "import os; print(os.environ.get('CUSTOM_VAR'))"],
    capture_output=True,
    text=True,
    env=env
)
print(result.stdout.strip())
"#;
    assert_eq!(run_python(src), vec!["custom_val_123"]);
}

#[test]
fn test_py_subprocess_devnull_redirection() {
    let src = r#"
import subprocess

result = subprocess.run(
    ["python3", "-c", "print('ignored output')"],
    stdout=subprocess.DEVNULL
)
print(result.returncode)
"#;
    assert_eq!(run_python(src), vec!["0"]);
}

#[test]
fn test_py_subprocess_shell_execution() {
    let src = r#"
import subprocess

result = subprocess.run("echo $((5 + 5))", shell=True, capture_output=True, text=True)
print(result.stdout.strip())
"#;
    assert_eq!(run_python(src), vec!["10"]);
}

#[test]
fn test_py_subprocess_popen_poll_terminate() {
    let src = r#"
import subprocess, time

proc = subprocess.Popen(["python3", "-c", "import time; time.sleep(10)"])
print(proc.poll() is None)  # running
proc.terminate()
proc.wait()
print(proc.poll() is not None)  # terminated
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_subprocess_cwd_working_directory() {
    let src = r#"
import subprocess, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    result = subprocess.run(
        ["python3", "-c", "import os; print(os.getcwd())"],
        cwd=tmpdir,
        capture_output=True,
        text=True
    )
    print(result.stdout.strip() == tmpdir)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}
