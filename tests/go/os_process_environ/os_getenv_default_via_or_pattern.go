// vybe-test: go/os_process_environ/os_getenv_default_via_or_pattern
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { v := os.Getenv("VYBE_OR_DEFAULT")
if v == "" { v = "default" }
_ = v }
