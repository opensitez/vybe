// vybe-test: go/os_exec_compile/getenv_concatenated_with_literal
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "os"
func main() { _ = os.Getenv("PREFIX") + "_suffix" }
