// vybe-test: go/os_process_environ/os_unsetenv_after_set
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { os.Setenv("VYBE_UNSET_TEST", "1")
os.Unsetenv("VYBE_UNSET_TEST")
_ = os.Getenv("VYBE_UNSET_TEST") }
