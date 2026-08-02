// vybe-test: go/struct_tags_reflect/reflect_tag_get_json_returns_value_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
import "reflect"
type Item struct { Count int `json:"count,omitempty"` }
func main() { _ = reflect.TypeOf(Item{}).Field(0).Tag.Get("json") }
