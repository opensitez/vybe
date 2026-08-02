// vybe-test: go/os_process_environ/os_environ_lookup_prefix
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
import "strings"
func main() { for _, e := range os.Environ() { if strings.HasPrefix(e, "PATH=") { _ = e } } }
