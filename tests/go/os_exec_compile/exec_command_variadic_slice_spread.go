// vybe-test: go/os_exec_compile/exec_command_variadic_slice_spread
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "os/exec"
func main() { flags := []string{"-n"}
_ = exec.Command("wc", flags...) }
