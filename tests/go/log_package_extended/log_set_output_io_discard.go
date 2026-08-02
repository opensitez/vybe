// vybe-test: go/log_package_extended/log_set_output_io_discard
// origin: languages/go/tests/go/test_log_package_extended.rs
// vybe-test-mode: compile

package main
import "log"
import "io"
func main() { log.SetOutput(io.Discard)
log.Print("gone") }
