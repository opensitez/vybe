// vybe-test: go/os_exec_compile/os_args_copy_into_slice
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "os"
func main() { copied := make([]string, len(os.Args))
copy(copied, os.Args) }
