// vybe-test: go/log_flag_packages/flag_bool_var_package_scope_compile
// origin: languages/go/tests/go/test_log_flag_packages.rs
// vybe-test-mode: compile

package main
import "flag"
var dryRun = flag.Bool("dry-run", false, "skip writes")
func main() { flag.Parse()
_ = *dryRun }
