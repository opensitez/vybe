// vybe-test: go/function_literals_closures/recursive_closure_via_var_binding
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var fib func(int) int
fib = func(n int) int { if n < 2 { return n }
return fib(n-1) + fib(n-2) }
__check(fmt.Sprint(fib(6)), "8") }
