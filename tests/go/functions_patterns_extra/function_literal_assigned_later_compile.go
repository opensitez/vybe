// vybe-test: go/functions_patterns_extra/function_literal_assigned_later_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { var fn func(int) int
fn = func(v int) int { return v }
_ = fn }
