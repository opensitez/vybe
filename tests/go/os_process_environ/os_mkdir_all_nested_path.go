// vybe-test: go/os_process_environ/os_mkdir_all_nested_path
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { _ = os.MkdirAll("/tmp/vybe/nested/dir", 0755) }
