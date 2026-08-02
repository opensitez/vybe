// vybe-test: go/builtins_expressions_extra/append_after_copy_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { dst := []int{0, 0}
copy(dst, []int{1, 2})
dst = append(dst, 3)
__check(fmt.Sprint(dst[2]), "3")
}
