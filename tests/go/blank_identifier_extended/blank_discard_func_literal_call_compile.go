// vybe-test: go/blank_identifier_extended/blank_discard_func_literal_call_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
func main() { _ = func(x int) int { return x }(4) }
