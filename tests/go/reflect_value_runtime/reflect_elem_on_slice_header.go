// vybe-test: go/reflect_value_runtime/reflect_elem_on_slice_header
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
func main() { s := []int{1}
_ = reflect.ValueOf(s).Index(0).Interface() }
