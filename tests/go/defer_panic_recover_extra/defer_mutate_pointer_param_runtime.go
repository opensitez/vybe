// vybe-test: go/defer_panic_recover_extra/defer_mutate_pointer_param_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := 1
func() { defer func(ptr *int) { *ptr = 5 }(&value) }()
__check(fmt.Sprint(value), "5")
}
