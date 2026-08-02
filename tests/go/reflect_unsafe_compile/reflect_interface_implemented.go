// vybe-test: go/reflect_unsafe_compile/reflect_interface_implemented
// origin: languages/go/tests/go/test_reflect_unsafe_compile.rs
// vybe-test-mode: compile

package main
import "reflect"
type R interface { M() }
type T struct{}
func (T) M() {}
func main() { _ = reflect.TypeOf((*T)(nil)).Implements(reflect.TypeOf((*R)(nil)).Elem()) }
