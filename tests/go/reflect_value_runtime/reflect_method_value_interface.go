// vybe-test: go/reflect_value_runtime/reflect_method_value_interface
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
type R struct{}
func (R) M() string { return "ok" }
func main() { _ = reflect.ValueOf(R{}).Method(0).Interface() }
