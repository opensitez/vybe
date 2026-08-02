// vybe-test: go/flag_parsing_extended/flag_panic_on_error_constant
// origin: languages/go/tests/go/test_flag_parsing_extended.rs
// vybe-test-mode: compile

package main
import "flag"
func main() { _ = flag.PanicOnError }
