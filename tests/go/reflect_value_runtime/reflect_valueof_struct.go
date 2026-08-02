// vybe-test: go/reflect_value_runtime/reflect_valueof_struct
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
type S struct { X int }
func main() { _ = reflect.ValueOf(S{X: 1}).Field(0) }
