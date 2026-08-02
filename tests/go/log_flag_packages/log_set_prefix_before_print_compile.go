// vybe-test: go/log_flag_packages/log_set_prefix_before_print_compile
// origin: languages/go/tests/go/test_log_flag_packages.rs
// vybe-test-mode: compile

package main
import "log"
func main() { log.SetPrefix("[app] ")
log.Print("ready") }
