// vybe-test: go/log_package_extended/log_set_flags_std_flags
// origin: languages/go/tests/go/test_log_package_extended.rs
// vybe-test-mode: compile

package main
import "log"
func main() { log.SetFlags(log.LstdFlags)
log.Print("std") }
