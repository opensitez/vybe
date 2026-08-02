// vybe-test: go/function_literals_closures/closure_passed_as_argument
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func apply(fn func(int) int, x int) int { return fn(x) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { sq := func(n int) int { return n * n }
__check(fmt.Sprint(apply(sq, 6)), "36") }
