use super::helpers::run_python;

// subprocess — run, Popen, communicate, check_output, check_call, PIPE, STDOUT, DEVNULL, CalledProcessError, TimeoutExpired

#[test]
fn test_subprocess_run_echo_stdout() {
    let out = run_python(r#"
import subprocess, sys
res = subprocess.run([sys.executable, "-c", "print('hello subprocess')"], capture_output=True, text=True)
print(res.returncode)
print(res.stdout.strip())
"#);
    assert_eq!(out, vec!["0", "hello subprocess"]);
}

#[test]
fn test_subprocess_run_check_raises_called_process_error() {
    let out = run_python(r#"
import subprocess, sys
try:
    subprocess.run([sys.executable, "-c", "import sys; sys.exit(42)"], check=True)
except subprocess.CalledProcessError as e:
    print(e.returncode)
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_subprocess_check_output_returns_bytes() {
    let out = run_python(r#"
import subprocess, sys
out = subprocess.check_output([sys.executable, "-c", "print('output text')"])
print(isinstance(out, bytes))
print(b"output text" in out)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_subprocess_check_output_text_mode() {
    let out = run_python(r#"
import subprocess, sys
out = subprocess.check_output([sys.executable, "-c", "print('text output')"], text=True)
print(isinstance(out, str))
print(out.strip())
"#);
    assert_eq!(out, vec!["True", "text output"]);
}

#[test]
fn test_subprocess_popen_communicate_stdin() {
    let out = run_python(r#"
import subprocess, sys
code = "import sys; print(sys.stdin.read().upper())"
proc = subprocess.Popen([sys.executable, "-c", code], stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True)
stdout, stderr = proc.communicate(input="input to stdin")
print(stdout.strip())
"#);
    assert_eq!(out, vec!["INPUT TO STDIN"]);
}

#[test]
fn test_subprocess_popen_poll_and_wait() {
    let out = run_python(r#"
import subprocess, sys
proc = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(0.01)"])
print(proc.poll() is None or proc.poll() == 0)
ret = proc.wait()
print(ret)
"#);
    assert_eq!(out, vec!["True", "0"]);
}

#[test]
fn test_subprocess_stdout_stderr_redirect_stderr_to_stdout() {
    let out = run_python(r#"
import subprocess, sys
code = "import sys; sys.stdout.write('out\\n'); sys.stderr.write('err\\n')"
proc = subprocess.Popen([sys.executable, "-c", code], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
stdout, _ = proc.communicate()
print("out" in stdout and "err" in stdout)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_subprocess_devnull_redirection() {
    let out = run_python(r#"
import subprocess, sys
res = subprocess.run([sys.executable, "-c", "print('should be suppressed')"], stdout=subprocess.DEVNULL)
print(res.returncode)
print(res.stdout is None)
"#);
    assert_eq!(out, vec!["0", "True"]);
}

#[test]
fn test_subprocess_timeout_expired_exception() {
    let out = run_python(r#"
import subprocess, sys
try:
    subprocess.run([sys.executable, "-c", "import time; time.sleep(2)"], timeout=0.05)
except subprocess.TimeoutExpired as e:
    print("TimeoutExpired")
"#);
    assert_eq!(out, vec!["TimeoutExpired"]);
}

#[test]
fn test_subprocess_check_call_success() {
    let out = run_python(r#"
import subprocess, sys
ret = subprocess.check_call([sys.executable, "-c", "import sys; sys.exit(0)"])
print(ret)
"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_subprocess_completed_process_args() {
    let out = run_python(r#"
import subprocess, sys
cmd = [sys.executable, "-c", "pass"]
res = subprocess.run(cmd)
print(res.args == cmd)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_subprocess_popen_pid_attribute() {
    let out = run_python(r#"
import subprocess, sys
proc = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(0.01)"])
print(isinstance(proc.pid, int))
print(proc.pid > 0)
proc.wait()
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_subprocess_popen_terminate_and_kill() {
    let out = run_python(r#"
import subprocess, sys
proc = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(10)"])
proc.terminate()
proc.wait()
print(proc.returncode != 0)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_subprocess_run_cwd_parameter() {
    let out = run_python(r#"
import subprocess, sys, tempfile
with tempfile.TemporaryDirectory() as tmpdir:
    res = subprocess.run([sys.executable, "-c", "import os; print(os.getcwd())"], cwd=tmpdir, capture_output=True, text=True)
    print(res.stdout.strip() == tmpdir or os.path.samefile(res.stdout.strip(), tmpdir))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_subprocess_run_env_parameter() {
    let out = run_python(r#"
import subprocess, sys, os
custom_env = dict(os.environ, MY_CUSTOM_VAR="vybe_test_val")
res = subprocess.run([sys.executable, "-c", "import os; print(os.environ.get('MY_CUSTOM_VAR'))"], env=custom_env, capture_output=True, text=True)
print(res.stdout.strip())
"#);
    assert_eq!(out, vec!["vybe_test_val"]);
}

#[test]
fn test_subprocess_popen_context_manager() {
    let out = run_python(r#"
import subprocess, sys
with subprocess.Popen([sys.executable, "-c", "print('inside popen')"], stdout=subprocess.PIPE, text=True) as proc:
    out, _ = proc.communicate()
    print(out.strip())
"#);
    assert_eq!(out, vec!["inside popen"]);
}

#[test]
fn test_subprocess_list2cmdline_windows_helper() {
    let out = run_python(r#"
import subprocess
s = subprocess.list2cmdline(["python", "script.py", "arg with spaces"])
print("arg with spaces" in s or '"arg with spaces"' in s)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_subprocess_run_input_parameter() {
    let out = run_python(r#"
import subprocess, sys
res = subprocess.run([sys.executable, "-c", "import sys; print(sys.stdin.read().strip())"], input="hello run input", text=True, capture_output=True)
print(res.stdout.strip())
"#);
    assert_eq!(out, vec!["hello run input"]);
}

#[test]
fn test_subprocess_pipe_constant() {
    let out = run_python(r#"
import subprocess
print(subprocess.PIPE == -1)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_subprocess_called_process_error_attributes() {
    let out = run_python(r#"
import subprocess
err = subprocess.CalledProcessError(1, ["cmd"], output="out", stderr="err")
print(err.returncode)
print(err.cmd)
print(err.output)
print(err.stderr)
"#);
    assert_eq!(out, vec!["1", "['cmd']", "out", "err"]);
}
