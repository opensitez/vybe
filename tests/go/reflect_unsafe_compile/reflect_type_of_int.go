// vybe-test: go/reflect_unsafe_compile/reflect_type_of_int
// origin: languages/go/tests/go/test_reflect_unsafe_compile.rs
// vybe-test-mode: compile

package main
import "reflect"
func main() { _ = reflect.TypeOf(0) }
