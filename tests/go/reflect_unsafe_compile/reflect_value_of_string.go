// vybe-test: go/reflect_unsafe_compile/reflect_value_of_string
// origin: languages/go/tests/go/test_reflect_unsafe_compile.rs
// vybe-test-mode: compile

package main
import "reflect"
func main() { _ = reflect.ValueOf("x") }
