// vybe-test: go/function_types_advanced/slice_of_named_func_type_compile
// origin: languages/go/tests/go/test_function_types_advanced.rs
// vybe-test-mode: compile

package main
type Op func(int) int
func main() { ops := []Op{Op(func(v int) int { return v }), Op(func(v int) int { return v + 1 })}
_ = ops[1](2) }
