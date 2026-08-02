// vybe-test: go/log_flag_packages/flag_string_var_package_scope_compile
// origin: languages/go/tests/go/test_log_flag_packages.rs
// vybe-test-mode: compile

package main
import "flag"
var region = flag.String("region", "us", "region code")
func main() { flag.Parse()
_ = *region }
