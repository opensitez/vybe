// vybe-test: go/os_process_environ/os_mkdir_temp_in_current_dir
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { dir, err := os.MkdirTemp(".", "local-*")
if err == nil { defer os.RemoveAll(dir) }
_ = dir }
