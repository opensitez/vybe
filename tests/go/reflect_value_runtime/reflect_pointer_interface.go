// vybe-test: go/reflect_value_runtime/reflect_pointer_interface
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
func main() { x := 1
_ = reflect.ValueOf(&x).Interface().(*int) }
