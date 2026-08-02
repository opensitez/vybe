// vybe-test: go/log_flag_packages/flag_parse_in_init_compile
// origin: languages/go/tests/go/test_log_flag_packages.rs
// vybe-test-mode: compile

package main
import "flag"
func init() { flag.Parse() }
func main() {}
