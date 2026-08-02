// vybe-test: go/log_package_extended/log_panic_two_args
// origin: languages/go/tests/go/test_log_package_extended.rs
// vybe-test-mode: compile

package main
import "log"
func main() { log.Panic("fail", 9) }
