// vybe-test: go/os_process_environ/os_mkdir_temp_with_pattern
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { dir, err := os.MkdirTemp("", "vybe-test-*")
if err == nil { defer os.RemoveAll(dir) }
_ = dir }
