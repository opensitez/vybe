// vybe-test: go/defer_panic_recover_extra/defer_in_branch_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { if true { defer __check(fmt.Sprint(2), "1") }
__check(fmt.Sprint(1), "2")
}
