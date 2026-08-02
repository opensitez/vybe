// vybe-test: go/os_exec_compile/getenv_lookup_assign
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "os"
func main() { v := os.Getenv("HOME")
_ = v }
