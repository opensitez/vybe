// vybe-test: go/reflect_unsafe_compile/reflect_ptr_elem
// origin: languages/go/tests/go/test_reflect_unsafe_compile.rs
// vybe-test-mode: compile

package main
import "reflect"
func main() { var x int
_ = reflect.TypeOf(&x).Elem() }
