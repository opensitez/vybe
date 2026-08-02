// vybe-test: go/os_exec_compile/os_args_len_used_in_make
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "os"
func main() { buf := make([]string, len(os.Args))
_ = buf }
