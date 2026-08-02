// vybe-test: go/reflect_value_runtime/reflect_value_convert_int_to_int64
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
func main() { v := reflect.ValueOf(int(5))
_ = v.Convert(reflect.TypeOf(int64(0))) }
