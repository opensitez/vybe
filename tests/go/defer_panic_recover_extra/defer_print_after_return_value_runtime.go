// vybe-test: go/defer_panic_recover_extra/defer_print_after_return_value_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func build() int { defer __check(fmt.Sprint("later"), "later")
return 4 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(build()), "4")
}
