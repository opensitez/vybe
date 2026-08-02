// vybe-test: go/log_flag_packages/log_flags_date_and_time_compile
// origin: languages/go/tests/go/test_log_flag_packages.rs
// vybe-test-mode: compile

package main
import "log"
func main() { log.SetFlags(log.Ldate | log.Ltime)
log.Print("stamp") }
