// vybe-test: go/reflect_unsafe_compile/reflect_struct_field_count
// origin: languages/go/tests/go/test_reflect_unsafe_compile.rs
// vybe-test-mode: compile

package main
import "reflect"
type S struct { A int
B string }
func main() { _ = reflect.TypeOf(S{}).NumField() }
