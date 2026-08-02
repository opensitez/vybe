// vybe-test: go/reflect_value_runtime/reflect_callslice_variadic
// origin: languages/go/tests/go/test_reflect_value_runtime.rs

package main
import "fmt"
import "reflect"
func Sum(nums ...int) int { s := 0
for _, n := range nums { s += n }
return s }
func main() { fv := reflect.ValueOf(Sum)
out := fv.CallSlice([]reflect.Value{reflect.ValueOf(1), reflect.ValueOf(2), reflect.ValueOf(3)})
fmt.Println(out[0].Int()) }
