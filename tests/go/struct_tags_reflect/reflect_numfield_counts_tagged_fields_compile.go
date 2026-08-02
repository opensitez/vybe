// vybe-test: go/struct_tags_reflect/reflect_numfield_counts_tagged_fields_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
import "reflect"
type Row struct { ID int `json:"id"`
Name string `json:"name"` }
func main() { _ = reflect.TypeOf(Row{}).NumField() }
