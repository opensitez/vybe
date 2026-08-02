// vybe-test: go/function_literals_closures/literal_iife_returns_value
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { n := func() int { return 99 }()
__check(fmt.Sprint(n), "99") }
