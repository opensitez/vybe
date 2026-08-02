// vybe-test: go/defer_panic_recover_extra/defer_before_multiple_returns_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func build(flag bool) int { defer __check(fmt.Sprint("done"), "done")
if flag { return 1 }
return 2 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(build(false)), "2")
}
