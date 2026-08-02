// vybe-test: go/function_literals_closures/literal_call_immediately_iife
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { result := func(a int, b int) int { return a + b }(10, 20)
__check(fmt.Sprint(result), "30") }
