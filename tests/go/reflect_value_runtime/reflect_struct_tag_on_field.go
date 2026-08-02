// vybe-test: go/reflect_value_runtime/reflect_struct_tag_on_field
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
type T struct { X int `json:"x"` }
func main() { _ = reflect.TypeOf(T{}).Field(0).Tag.Get("json") }
