// vybe-test: go/os_process_environ/os_getenv_in_string_concat
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { _ = "prefix_" + os.Getenv("HOME") }
