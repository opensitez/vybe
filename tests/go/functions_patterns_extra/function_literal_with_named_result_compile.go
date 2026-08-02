// vybe-test: go/functions_patterns_extra/function_literal_with_named_result_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { fn := func(v int) (result int) { result = v + 1
return }
_ = fn }
