// vybe-test: go/struct_tags_reflect/reflect_field_by_name_found_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
import "reflect"
type Row struct { Score int `json:"score"` }
func main() { _, _ = reflect.TypeOf(Row{}).FieldByName("Score") }
