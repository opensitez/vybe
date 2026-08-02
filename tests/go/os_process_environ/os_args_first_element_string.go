// vybe-test: go/os_process_environ/os_args_first_element_string
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { if len(os.Args) > 0 { _ = os.Args[0] } }
