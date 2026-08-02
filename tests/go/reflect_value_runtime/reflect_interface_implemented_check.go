// vybe-test: go/reflect_value_runtime/reflect_interface_implemented_check
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
type I interface { F() }
type T struct{}
func (T) F() {}
func main() { _ = reflect.TypeOf(T{}).Implements(reflect.TypeOf((*I)(nil)).Elem()) }
