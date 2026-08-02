// vybe-test: go/log_package_extended/log_set_output_os_stderr
// origin: languages/go/tests/go/test_log_package_extended.rs
// vybe-test-mode: compile

package main
import "log"
import "os"
func main() { log.SetOutput(os.Stderr)
log.Print("e") }
