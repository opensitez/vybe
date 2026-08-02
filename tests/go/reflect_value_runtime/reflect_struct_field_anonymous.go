// vybe-test: go/reflect_value_runtime/reflect_struct_field_anonymous
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
type Inner struct { N int }
type Outer struct { Inner }
func main() { _ = reflect.TypeOf(Outer{}).Field(0) }
