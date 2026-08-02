// vybe-test: go/function_literals_closures/literal_capture_multiple_outer_vars
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := 2
b := 3
sum := func() int { return a + b }
__check(fmt.Sprint(sum()), "5") }
