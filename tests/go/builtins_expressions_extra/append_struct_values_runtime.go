// vybe-test: go/builtins_expressions_extra/append_struct_values_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

package main
import "fmt"
type point struct { x int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := []point{}
values = append(values, point{x: 14})
__check(fmt.Sprint(values[0].x), "14")
}
