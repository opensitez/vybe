// vybe-test: go/reflect_value_runtime/reflect_slice_append_reflect
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
func main() { s := []int{1}
sv := reflect.ValueOf(&s).Elem()
sv.Set(reflect.Append(sv, reflect.ValueOf(2))) }
