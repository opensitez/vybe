// vybe-test: go/reflect_value_runtime/reflect_set_uint_on_elem
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
func main() { var x uint
v := reflect.ValueOf(&x).Elem()
v.SetUint(7) }
