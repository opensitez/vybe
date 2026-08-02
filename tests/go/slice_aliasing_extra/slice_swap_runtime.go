// vybe-test: go/slice_aliasing_extra/slice_swap_runtime
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := []int{1, 2}
values[0], values[1] = values[1], values[0]
__check(fmt.Sprint(values[0]), "2")
__check(fmt.Sprint(values[1]), "1")
}
