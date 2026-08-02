// vybe-test: go/log_flag_packages/log_fatal_with_code_compile
// origin: languages/go/tests/go/test_log_flag_packages.rs
// vybe-test-mode: compile

package main
import "log"
func main() { log.Fatal("exit", 1) }
