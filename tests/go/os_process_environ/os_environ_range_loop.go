// vybe-test: go/os_process_environ/os_environ_range_loop
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { for _, entry := range os.Environ() { _ = entry } }
