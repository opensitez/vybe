// vybe-test: go/function_types_advanced/method_runs_supplied_transform_on_field
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type gauge struct { value int }
func (g *gauge) mapValue(mapper func(int) int) { g.value = mapper(g.value) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { g := gauge{value: 4}
g.mapValue(func(v int) int { return v * 2 })
__check(fmt.Sprint(g.value), "8") }
