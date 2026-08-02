// vybe-test: go/os_process_environ/os_pipe_create
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { r, w, err := os.Pipe()
if err == nil { r.Close()
w.Close() } }
