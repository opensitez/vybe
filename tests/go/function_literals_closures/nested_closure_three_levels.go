// vybe-test: go/function_literals_closures/nested_closure_three_levels
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { level1 := func(a int) func(b int) func(c int) int { return func(b int) func(c int) int { return func(c int) int { return a + b + c } } }
fn := level1(1)(2)
__check(fmt.Sprint(fn(3)), "6") }
