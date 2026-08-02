// vybe-test: go/log_flag_packages/flag_narg_after_parse_compile
// origin: languages/go/tests/go/test_log_flag_packages.rs
// vybe-test-mode: compile

package main
import "flag"
func main() { flag.Parse()
_ = flag.NArg() }
