// vybe-test: go/os_process_environ/os_args_index_bounds_check
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { for i := range os.Args { _ = os.Args[i] } }
