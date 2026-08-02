// vybe-test: go/struct_tags_reflect/reflect_field_tag_property_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
import "reflect"
type Row struct { Name string `json:"name"` }
func main() { _ = reflect.TypeOf(Row{}).Field(0).Tag }
