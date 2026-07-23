use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: subprocess + os.popen — running processes, capturing output, pipes, returncode, environment
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_subprocess_run_basic() {
    let src = r#"
import subprocess

result = subprocess.run(["echo", "hello"], capture_output=True, text=True)
print(result.returncode)
print(result.stdout.strip())
print(result.stderr)
"#;
    assert_eq!(run_python(src), vec!["0", "hello", ""]);
}

#[test]
fn test_py_subprocess_run_with_input() {
    let src = r#"
import subprocess

result = subprocess.run(
    ["cat"],
    input="test input",
    capture_output=True,
    text=True
)
print(result.returncode)
print(result.stdout)
"#;
    assert_eq!(run_python(src), vec!["0", "test input"]);
}

#[test]
fn test_py_subprocess_shell_true() {
    let src = r#"
import subprocess

result = subprocess.run("echo $((2 + 2))", shell=True, capture_output=True, text=True)
print(result.returncode)
print(result.stdout.strip())
"#;
    assert_eq!(run_python(src), vec!["0", "4"]);
}

#[test]
fn test_py_subprocess_non_zero_exit_code() {
    let src = r#"
import subprocess

result = subprocess.run(["false"], capture_output=True)
print(result.returncode != 0)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_subprocess_check_raises_on_error() {
    let src = r#"
import subprocess

try:
    subprocess.run(["false"], check=True)
except subprocess.CalledProcessError as e:
    print(f"CalledProcessError: returncode={e.returncode}")
"#;
    assert_eq!(run_python(src), vec!["CalledProcessError: returncode=1"]);
}

#[test]
fn test_py_subprocess_check_output() {
    let src = r#"
import subprocess

output = subprocess.check_output(["echo", "hello world"], text=True)
print(output.strip())
"#;
    assert_eq!(run_python(src), vec!["hello world"]);
}

#[test]
fn test_py_subprocess_timeout() {
    let src = r#"
import subprocess

try:
    subprocess.run(["sleep", "10"], timeout=0.1)
except subprocess.TimeoutExpired:
    print("TimeoutExpired")
"#;
    assert_eq!(run_python(src), vec!["TimeoutExpired"]);
}

#[test]
fn test_py_subprocess_popen_pipe() {
    let src = r#"
import subprocess

proc = subprocess.Popen(
    ["python3", "-c", "import sys; sys.stdout.write('from_child')"],
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE
)
stdout, stderr = proc.communicate()
print(stdout.decode())
print(proc.returncode)
"#;
    assert_eq!(run_python(src), vec!["from_child", "0"]);
}

#[test]
fn test_py_subprocess_env_override() {
    let src = r#"
import subprocess, os

env = os.environ.copy()
env["MY_VAR"] = "custom_value"

result = subprocess.run(
    ["python3", "-c", "import os; print(os.environ.get('MY_VAR', 'not set'))"],
    capture_output=True,
    text=True,
    env=env
)
print(result.stdout.strip())
"#;
    assert_eq!(run_python(src), vec!["custom_value"]);
}

#[test]
fn test_py_subprocess_cwd_change() {
    let src = r#"
import subprocess, tempfile, os

with tempfile.TemporaryDirectory() as tmpdir:
    result = subprocess.run(
        ["python3", "-c", "import os; print(os.getcwd())"],
        capture_output=True,
        text=True,
        cwd=tmpdir
    )
    print(result.stdout.strip() == tmpdir)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}
