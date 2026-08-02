// vybe-test: go/log_flag_packages/log_print_inside_defer_compile
// origin: languages/go/tests/go/test_log_flag_packages.rs
// vybe-test-mode: compile

package main
import "log"
func main() { defer log.Print("bye")
log.Print("hi") }
