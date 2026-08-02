// vybe-test: go/os_process_environ/os_create_temp_file
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { f, err := os.CreateTemp("", "vybe-*")
if err == nil { defer os.Remove(f.Name())
defer f.Close() } }
