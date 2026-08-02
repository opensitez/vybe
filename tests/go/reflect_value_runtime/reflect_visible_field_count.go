// vybe-test: go/reflect_value_runtime/reflect_visible_field_count
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
type S struct { Pub int
priv int }
func main() { _ = reflect.TypeOf(S{}).NumField() }
