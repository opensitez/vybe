// vybe-test: go/os_process_environ/os_open_file_readonly
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { f, err := os.Open(".")
if err == nil { defer f.Close() } }
