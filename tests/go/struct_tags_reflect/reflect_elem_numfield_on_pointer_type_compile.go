// vybe-test: go/struct_tags_reflect/reflect_elem_numfield_on_pointer_type_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
import "reflect"
type Pair struct { A int `json:"a"`
B int `json:"b"`
C int `json:"c"` }
func main() { _ = reflect.TypeOf(&Pair{}).Elem().NumField() }
