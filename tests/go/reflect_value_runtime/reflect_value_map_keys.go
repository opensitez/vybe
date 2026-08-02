// vybe-test: go/reflect_value_runtime/reflect_value_map_keys
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
func main() { _ = reflect.ValueOf(map[int]int{1: 1}).MapKeys() }
