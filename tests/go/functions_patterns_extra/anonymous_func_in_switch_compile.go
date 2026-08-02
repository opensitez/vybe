// vybe-test: go/functions_patterns_extra/anonymous_func_in_switch_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { switch func() int { return 2 }() { case 2: _ = 2 } }
