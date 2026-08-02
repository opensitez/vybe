// vybe-test: go/log_package_extended/log_print_to_bytes_buffer_var
// origin: languages/go/tests/go/test_log_package_extended.rs
// vybe-test-mode: compile

package main
import "log"
import "bytes"
func main() { var buf bytes.Buffer
log.SetOutput(&buf)
log.SetFlags(0)
log.Print("b") }
