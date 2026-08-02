// vybe-test: go/os_process_environ/os_open_file_with_flags
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { _, _ = os.OpenFile("/tmp/vybe-open", os.O_CREATE|os.O_RDWR, 0644) }
