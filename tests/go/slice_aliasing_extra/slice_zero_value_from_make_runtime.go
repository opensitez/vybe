// vybe-test: go/slice_aliasing_extra/slice_zero_value_from_make_runtime
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := make([]int, 2)
__check(fmt.Sprint(values[0]), "0")
__check(fmt.Sprint(values[1]), "0")
}
