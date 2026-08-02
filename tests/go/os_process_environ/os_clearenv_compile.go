// vybe-test: go/os_process_environ/os_clearenv_compile
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { os.Clearenv()
_ = len(os.Environ()) }
