// vybe-test: go/slice_aliasing_extra/copy_slice_returns_count_runtime
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { dst := make([]int, 2)
src := []int{3, 4, 5}
__check(fmt.Sprint(copy(dst, src)), "2")
}
