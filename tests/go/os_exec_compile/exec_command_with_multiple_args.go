// vybe-test: go/os_exec_compile/exec_command_with_multiple_args
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "os/exec"
func main() { _ = exec.Command("sh", "-c", "echo hi") }
