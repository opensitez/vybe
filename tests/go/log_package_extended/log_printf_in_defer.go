// vybe-test: go/log_package_extended/log_printf_in_defer
// origin: languages/go/tests/go/test_log_package_extended.rs
// vybe-test-mode: compile

package main
import "log"
func main() { defer log.Printf("d=%d", 1) }
