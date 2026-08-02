// vybe-test: go/function_literals_closures/closure_returns_closure
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func outer(x int) func(int) func(int) int { return func(y int) func(int) int { return func(z int) int { return x + y + z } } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { fn := outer(1)(2)
__check(fmt.Sprint(fn(3)), "6") }
