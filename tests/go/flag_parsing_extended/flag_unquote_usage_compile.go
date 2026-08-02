// vybe-test: go/flag_parsing_extended/flag_unquote_usage_compile
// origin: languages/go/tests/go/test_flag_parsing_extended.rs
// vybe-test-mode: compile

package main
import "flag"
func main() { _ = flag.UnquoteUsage }
