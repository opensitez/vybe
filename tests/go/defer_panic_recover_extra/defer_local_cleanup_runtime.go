// vybe-test: go/defer_panic_recover_extra/defer_local_cleanup_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { done := 0
func() { defer func() { done = 3 }() }()
__check(fmt.Sprint(done), "3")
}
