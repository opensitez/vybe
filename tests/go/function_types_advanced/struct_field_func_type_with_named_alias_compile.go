// vybe-test: go/function_types_advanced/struct_field_func_type_with_named_alias_compile
// origin: languages/go/tests/go/test_function_types_advanced.rs
// vybe-test-mode: compile

package main
type Callback func(int) bool
type registry struct { filter Callback }
func main() { _ = registry{filter: Callback(func(v int) bool { return v > 0 })} }
