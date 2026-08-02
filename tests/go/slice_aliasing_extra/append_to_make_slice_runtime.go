// vybe-test: go/slice_aliasing_extra/append_to_make_slice_runtime
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := make([]int, 0, 3)
values = append(values, 6)
__check(fmt.Sprint(values[0]), "6")
}
