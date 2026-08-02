// vybe-test: go/log_package_extended/log_set_flags_all_metadata
// origin: languages/go/tests/go/test_log_package_extended.rs
// vybe-test-mode: compile

package main
import "log"
func main() { log.SetFlags(log.Ldate | log.Ltime | log.Lmicroseconds | log.Llongfile)
log.Print("all") }
