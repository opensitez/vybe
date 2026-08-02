// vybe-test: go/builtins_expressions_extra/len_on_nested_slice_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { grid := [][]int{{1}, {2}, {3}}
__check(fmt.Sprint(len(grid)), "3")
}
