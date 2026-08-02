// vybe-test: go/os_process_environ/os_expand_with_mapping_func
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { _ = os.Expand("$USER", func(k string) string { return os.Getenv(k) }) }
