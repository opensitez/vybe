// vybe-test: go/slice_aliasing_extra/slice_append_preserves_prefix_runtime
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := []int{2, 4}
values = append(values, 6)
__check(fmt.Sprint(values[0]), "2")
__check(fmt.Sprint(values[2]), "6")
}
