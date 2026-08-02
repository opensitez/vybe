// vybe-test: go/reflect_value_runtime/reflect_typeof_chan
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
func main() { _ = reflect.TypeOf(make(chan int)).Kind() }
