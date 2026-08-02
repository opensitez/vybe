// vybe-test: go/log_flag_packages/flag_int_var_package_scope_compile
// origin: languages/go/tests/go/test_log_flag_packages.rs
// vybe-test-mode: compile

package main
import "flag"
var workers = flag.Int("workers", 4, "worker count")
func main() { flag.Parse()
_ = *workers }
