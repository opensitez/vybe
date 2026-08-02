// vybe-test: go/slice_aliasing_extra/append_slice_expansion_runtime
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
values = append(values, 3, 4)
__check(fmt.Sprint(len(values)), "4")
__check(fmt.Sprint(values[3]), "4")
}
