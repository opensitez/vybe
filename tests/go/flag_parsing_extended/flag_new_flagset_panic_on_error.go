// vybe-test: go/flag_parsing_extended/flag_new_flagset_panic_on_error
// origin: languages/go/tests/go/test_flag_parsing_extended.rs
// vybe-test-mode: compile

package main
import "flag"
func main() { _ = flag.NewFlagSet("app", flag.PanicOnError) }
