// vybe-test: go/functions_patterns_extra/nested_function_literals_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { outer := func() func() int { return func() int { return 1 } }
_ = outer }
