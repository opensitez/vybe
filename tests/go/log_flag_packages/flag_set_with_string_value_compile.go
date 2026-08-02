// vybe-test: go/log_flag_packages/flag_set_with_string_value_compile
// origin: languages/go/tests/go/test_log_flag_packages.rs
// vybe-test-mode: compile

package main
import "flag"
func main() { f := flag.String("mode", "dev", "")
_ = flag.Set("mode", "prod")
_ = *f }
