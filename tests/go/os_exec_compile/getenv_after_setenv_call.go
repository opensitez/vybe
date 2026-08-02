// vybe-test: go/os_exec_compile/getenv_after_setenv_call
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "os"
func main() { os.Setenv("VYBE_TEST", "1")
_ = os.Getenv("VYBE_TEST") }
