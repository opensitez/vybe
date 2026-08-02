// vybe-test: go/reflect_value_runtime/reflect_indirect_on_pointer
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
func main() { x := 3
_ = reflect.Indirect(reflect.ValueOf(&x)).Int() }
