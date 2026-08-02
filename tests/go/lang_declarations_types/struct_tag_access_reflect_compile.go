// vybe-test: go/lang_declarations_types/struct_tag_access_reflect_compile
// origin: languages/go/tests/go/test_lang_declarations_types.rs
// vybe-test-mode: compile

package main
import "reflect"
type S struct { X int `json:"x"` }
func main() { _ = reflect.TypeOf(S{}).Field(0).Tag }
