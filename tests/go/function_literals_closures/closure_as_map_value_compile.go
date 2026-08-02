// vybe-test: go/function_literals_closures/closure_as_map_value_compile
// origin: languages/go/tests/go/test_function_literals_closures.rs
// vybe-test-mode: compile

package main
func main() { m := map[int]func(){ 1: func() {} }
m[1]() }
