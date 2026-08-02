// vybe-test: go/slice_aliasing_extra/slice_alias_after_assignment_runtime
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { left := []int{1, 2}
right := left
right[0] = 8
__check(fmt.Sprint(left[0]), "8")
}
