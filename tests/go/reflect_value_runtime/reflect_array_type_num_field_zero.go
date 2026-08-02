// vybe-test: go/reflect_value_runtime/reflect_array_type_num_field_zero
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
func main() { _ = reflect.TypeOf([3]int{}).NumField() }
