// vybe-test: go/function_literals_closures/return_closure_from_function
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func makeAdder(base int) func(int) int { return func(x int) int { return base + x } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { add5 := makeAdder(5)
__check(fmt.Sprint(add5(3)), "8") }
