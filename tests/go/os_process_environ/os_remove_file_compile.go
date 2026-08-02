// vybe-test: go/os_process_environ/os_remove_file_compile
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { _ = os.Remove("/tmp/vybe-write-test.txt") }
