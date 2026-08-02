// vybe-test: go/nil_zero_semantics_extra/nil_slice_copy_runtime
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var dst []int
src := []int{1, 2}
__check(fmt.Sprint(copy(dst, src)), "0")
}
