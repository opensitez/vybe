// vybe-test: go/os_exec_compile/getenv_in_boolean_condition
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "os"
func main() { if os.Getenv("DEBUG") != "" { _ = 1 } }
