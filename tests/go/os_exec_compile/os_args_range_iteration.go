// vybe-test: go/os_exec_compile/os_args_range_iteration
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "os"
func main() { for _, arg := range os.Args { _ = arg } }
