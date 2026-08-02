// vybe-test: go/functions_patterns_extra/function_returning_named_type_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
type score int
func build() score { return 5 }
func main() { _ = build() }
