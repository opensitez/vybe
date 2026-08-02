// vybe-test: go/struct_tags_reflect/reflect_struct_tag_get_second_key_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
import "reflect"
type Row struct { Title string `json:"title" xml:"title"` }
func main() { _ = reflect.TypeOf(Row{}).Field(0).Tag.Get("xml") }
