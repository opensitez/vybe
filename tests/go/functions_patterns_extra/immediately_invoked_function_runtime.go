// vybe-test: go/functions_patterns_extra/immediately_invoked_function_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := func(n int) int { return n * 2 }(6)
__check(fmt.Sprint(value), "12")
}
