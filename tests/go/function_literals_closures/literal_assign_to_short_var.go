// vybe-test: go/function_literals_closures/literal_assign_to_short_var
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { double := func(x int) int { return x * 2 }
__check(fmt.Sprint(double(7)), "14") }
