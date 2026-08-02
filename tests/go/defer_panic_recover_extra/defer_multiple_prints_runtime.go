// vybe-test: go/defer_panic_recover_extra/defer_multiple_prints_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer __check(fmt.Sprint("a"), "a")
defer __check(fmt.Sprint("b"), "b")
__check(fmt.Sprint("c"), "c")
}
