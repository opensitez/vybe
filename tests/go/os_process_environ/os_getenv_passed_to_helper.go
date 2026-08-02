// vybe-test: go/os_process_environ/os_getenv_passed_to_helper
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func lookup(k string) string { return os.Getenv(k) }
func main() { _ = lookup("USER") }
