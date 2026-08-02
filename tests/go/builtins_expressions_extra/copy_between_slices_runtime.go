// vybe-test: go/builtins_expressions_extra/copy_between_slices_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { dst := []int{0, 0, 0}
src := []int{7, 8}
copy(dst, src)
__check(fmt.Sprint(dst[0]), "7")
__check(fmt.Sprint(dst[1]), "8")
}
