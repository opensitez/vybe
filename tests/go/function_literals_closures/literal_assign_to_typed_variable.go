// vybe-test: go/function_literals_closures/literal_assign_to_typed_variable
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var fn func(int) int = func(x int) int { return x * 3 }
__check(fmt.Sprint(fn(4)), "12") }
