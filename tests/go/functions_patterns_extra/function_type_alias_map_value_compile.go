// vybe-test: go/functions_patterns_extra/function_type_alias_map_value_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
type reducer func(int, int) int
func main() { ops := map[string]reducer{}
_ = ops }
