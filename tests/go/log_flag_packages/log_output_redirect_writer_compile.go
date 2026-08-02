// vybe-test: go/log_flag_packages/log_output_redirect_writer_compile
// origin: languages/go/tests/go/test_log_flag_packages.rs
// vybe-test-mode: compile

package main
import "log"
import "os"
func main() { log.SetOutput(os.Stderr)
log.Print("err") }
