// vybe-test: go/function_literals_closures/closure_stored_in_slice
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { fns := []func(int) int{ func(x int) int { return x + 1 }, func(x int) int { return x + 2 } }
__check(fmt.Sprint(fns[0](5)), "6")
__check(fmt.Sprint(fns[1](5)), "7") }
