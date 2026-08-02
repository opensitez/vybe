// vybe-test: go/builtins_expressions_extra/new_struct_pointer_runtime
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

func main() { value := new(point)
value.x = 12
__check(fmt.Sprint(value.x), "12")
}
