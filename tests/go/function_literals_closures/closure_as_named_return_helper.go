// vybe-test: go/function_literals_closures/closure_as_named_return_helper
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func compute() int { transform := func(x int) int { return x + 10 }
return transform(5) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(compute()), "15") }
