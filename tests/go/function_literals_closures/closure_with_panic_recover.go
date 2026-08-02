// vybe-test: go/function_literals_closures/closure_with_panic_recover
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { safe := func() { defer func() { __check(fmt.Sprint(recover() != nil), "true") }()
panic("x") }
safe() }
