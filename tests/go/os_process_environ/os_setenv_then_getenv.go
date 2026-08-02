// vybe-test: go/os_process_environ/os_setenv_then_getenv
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { os.Setenv("VYBE_SET_TEST", "42")
_ = os.Getenv("VYBE_SET_TEST") }
