// vybe-test: go/switch_fallthrough_extended/fallthrough_as_first_statement_compile
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs
// vybe-test-mode: compile

package main
func main() { switch 1 { case 1: fallthrough } }
