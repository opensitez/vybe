// vybe-test: go/os_exec_compile/os_args_join_with_separator
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "os"
import "strings"
func main() { _ = strings.Join(os.Args, " ") }
