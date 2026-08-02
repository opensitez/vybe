// vybe-test: go/struct_tags_reflect/reflect_field_zero_reads_first_tag_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
import "reflect"
type Row struct { ID int `json:"id"` }
func main() { _ = reflect.TypeOf(Row{}).Field(0) }
