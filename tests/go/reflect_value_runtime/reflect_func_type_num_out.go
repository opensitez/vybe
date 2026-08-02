// vybe-test: go/reflect_value_runtime/reflect_func_type_num_out
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
func main() { t := reflect.TypeOf(func() (int, error) { return 0, nil })
_ = t.NumOut() }
