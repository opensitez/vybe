// vybe-test: go/struct_tags_reflect/reflect_struct_tag_len_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
import "reflect"
type Row struct { Code string `json:"code"` }
func main() { _ = reflect.TypeOf(Row{}).Field(0).Tag.Len() }
