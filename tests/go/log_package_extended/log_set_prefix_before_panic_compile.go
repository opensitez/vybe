// vybe-test: go/log_package_extended/log_set_prefix_before_panic_compile
// origin: languages/go/tests/go/test_log_package_extended.rs
// vybe-test-mode: compile

package main
import "log"
func main() { log.SetPrefix("X")
log.Panic("p") }
