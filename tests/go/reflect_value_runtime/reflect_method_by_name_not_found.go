// vybe-test: go/reflect_value_runtime/reflect_method_by_name_not_found
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
type T struct{}
func main() { _ = reflect.ValueOf(T{}).MethodByName("Missing").IsValid() }
