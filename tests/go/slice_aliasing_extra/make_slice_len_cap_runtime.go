// vybe-test: go/slice_aliasing_extra/make_slice_len_cap_runtime
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := make([]int, 2, 5)
__check(fmt.Sprint(len(values)), "2")
__check(fmt.Sprint(cap(values)), "5")
}
