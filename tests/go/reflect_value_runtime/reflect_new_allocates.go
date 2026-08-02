// vybe-test: go/reflect_value_runtime/reflect_new_allocates
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
func main() { p := reflect.New(reflect.TypeOf(0))
_ = p.Elem().SetInt(1) }
