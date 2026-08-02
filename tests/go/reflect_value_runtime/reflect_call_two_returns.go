// vybe-test: go/reflect_value_runtime/reflect_call_two_returns
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
func DivMod(a, b int) (int, int) { return a / b, a % b }
func main() { out := reflect.ValueOf(DivMod).Call([]reflect.Value{reflect.ValueOf(10), reflect.ValueOf(3)})
_, _ = out[0], out[1] }
