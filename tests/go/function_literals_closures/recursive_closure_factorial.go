// vybe-test: go/function_literals_closures/recursive_closure_factorial
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var fact func(int) int
fact = func(n int) int { if n <= 1 { return 1 }
return n * fact(n-1) }
__check(fmt.Sprint(fact(5)), "120") }
