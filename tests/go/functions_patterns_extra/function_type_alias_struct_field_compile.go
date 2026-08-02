// vybe-test: go/functions_patterns_extra/function_type_alias_struct_field_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
type reducer func(int) int
type holder struct { fn reducer }
func main() { _ = holder{} }
