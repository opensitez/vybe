// vybe-test: go/struct_tags_reflect/reflect_typeof_pointer_to_tagged_struct_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
import "reflect"
type Tagged struct { X int `json:"x"` }
func main() { _ = reflect.TypeOf(&Tagged{}) }
