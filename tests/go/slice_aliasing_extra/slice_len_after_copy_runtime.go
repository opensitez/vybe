// vybe-test: go/slice_aliasing_extra/slice_len_after_copy_runtime
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { dst := make([]int, 3)
copy(dst, []int{1, 2})
__check(fmt.Sprint(len(dst)), "3")
}
