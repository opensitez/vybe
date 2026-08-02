// vybe-test: go/function_types_advanced/nested_func_return_type_as_field_compile
// origin: languages/go/tests/go/test_function_types_advanced.rs
// vybe-test-mode: compile

package main
type Factory func() func() int
type holder struct { build Factory }
func main() { _ = holder{build: Factory(func() func() int { return func() int { return 1 } })} }
