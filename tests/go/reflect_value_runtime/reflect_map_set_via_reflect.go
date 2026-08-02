// vybe-test: go/reflect_value_runtime/reflect_map_set_via_reflect
// origin: languages/go/tests/go/test_reflect_value_runtime.rs
// vybe-test-mode: compile

package main
import "reflect"
func main() { m := map[string]int{}
mv := reflect.ValueOf(m)
mv.SetMapIndex(reflect.ValueOf("k"), reflect.ValueOf(1)) }
