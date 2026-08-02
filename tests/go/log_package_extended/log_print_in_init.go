// vybe-test: go/log_package_extended/log_print_in_init
// origin: languages/go/tests/go/test_log_package_extended.rs
// vybe-test-mode: compile

package main
import "log"
func init() { log.Print("init") }
func main() {}
