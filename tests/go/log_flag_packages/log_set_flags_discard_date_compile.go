// vybe-test: go/log_flag_packages/log_set_flags_discard_date_compile
// origin: languages/go/tests/go/test_log_flag_packages.rs
// vybe-test-mode: compile

package main
import "log"
func main() { log.SetFlags(0)
log.Print("plain") }
