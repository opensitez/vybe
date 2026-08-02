// vybe-test: go/os_exec_compile/exec_command_stored_in_variable
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "os/exec"
func main() { cmd := exec.Command("date")
_ = cmd }
