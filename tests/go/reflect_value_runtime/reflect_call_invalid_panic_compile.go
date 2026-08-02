// vybe-test: go/reflect_value_runtime/reflect_call_invalid_panic_compile
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
func noop() {}
func main() { _ = reflect.ValueOf(noop).Call(nil) }
