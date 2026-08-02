// vybe-test: go/os_process_environ/os_lstat_symlink_target
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { _, _ = os.Lstat(".") }
