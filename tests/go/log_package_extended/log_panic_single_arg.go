// vybe-test: go/log_package_extended/log_panic_single_arg
// origin: languages/go/tests/go/test_log_package_extended.rs
// vybe-test-mode: compile

package main
import "log"
func main() { log.Panic("boom") }
