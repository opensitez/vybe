// vybe-test: go/os_process_environ/os_write_file_compile
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { _ = os.WriteFile("/tmp/vybe-write-test.txt", []byte("x"), 0644) }
