// vybe-test: go/builtins_expressions_extra/make_slice_zero_length_nonzero_cap_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := make([]int, 0, 4)
__check(fmt.Sprint(len(values)), "0")
__check(fmt.Sprint(cap(values)), "4")
}
