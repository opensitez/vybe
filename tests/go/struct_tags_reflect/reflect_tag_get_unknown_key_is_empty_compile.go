// vybe-test: go/struct_tags_reflect/reflect_tag_get_unknown_key_is_empty_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
import "reflect"
type Item struct { Label string `json:"label"` }
func main() { _ = reflect.TypeOf(Item{}).Field(0).Tag.Get("xml") }
