// vybe-test: go/struct_tags_reflect/reflect_field_type_kind_int_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
import "reflect"
type Row struct { Count int `json:"count"` }
func main() { _ = reflect.TypeOf(Row{}).Field(0).Type.Kind() }
