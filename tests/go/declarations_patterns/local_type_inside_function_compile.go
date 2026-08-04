// vybe-test: go/declarations_patterns/local_type_inside_function_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
func main() { type local struct { value int }
v := local{value: 3}
_ = v }
