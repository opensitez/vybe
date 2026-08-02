// vybe-test: go/defer_panic_recover_extra/defer_named_return_add_two_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func build() (result int) { defer func() { result += 2 }()
return 3 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(build()), "5")
}
