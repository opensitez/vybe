// vybe-test: go/builtins_expressions_extra/cap_on_array_slice_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := [5]int{1, 2, 3, 4, 5}
part := values[1:3]
__check(fmt.Sprint(cap(part)), "4")
}
