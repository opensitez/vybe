// vybe-test: go/builtins_expressions_extra/cap_on_make_slice_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := make([]int, 2, 5)
__check(fmt.Sprint(cap(values)), "5")
}
