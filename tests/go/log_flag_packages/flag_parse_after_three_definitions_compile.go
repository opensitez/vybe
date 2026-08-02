// vybe-test: go/log_flag_packages/flag_parse_after_three_definitions_compile
// origin: languages/go/tests/go/test_log_flag_packages.rs
// vybe-test-mode: compile

package main
import "flag"
func main() { _ = flag.String("host", "", "")
_ = flag.Int("port", 0, "")
_ = flag.Bool("tls", false, "")
flag.Parse() }
