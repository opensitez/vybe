// vybe-test: go/defer_panic_recover_extra/defer_modify_slice_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := []int{1}
func() { defer func() { values[0] = 9 }() }()
__check(fmt.Sprint(values[0]), "9")
}
