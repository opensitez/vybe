// vybe-test: go/os_process_environ/os_rename_file_compile
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { _ = os.Rename("/tmp/a", "/tmp/b") }
