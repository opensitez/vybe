// vybe-test: go/os_exec_compile/os_args_append_spread_into_local
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "os"
func main() { local := append([]string{"prog"}, os.Args...) }
