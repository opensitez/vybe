// vybe-test: go/higher_order_functions/compose_two_functions
// origin: languages/go/tests/go/test_higher_order_functions.rs

package main
import "fmt"
func compose(f func(int) int, g func(int) int) func(int) int { return func(x int) int { return f(g(x)) } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { double := func(x int) int { return x * 2 }
inc := func(x int) int { return x + 1 }
h := compose(double, inc)
__check(fmt.Sprint(h(3)), "8") }
