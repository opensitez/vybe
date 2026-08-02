// vybe-test: go/builtins_expressions_extra/make_slice_with_length_and_capacity_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := make([]int, 3, 6)
__check(fmt.Sprint(len(values)), "3")
__check(fmt.Sprint(cap(values)), "6")
}
