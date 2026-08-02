// vybe-test: go/function_literals_closures/closure_reassign_variable
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { fn := func() int { return 1 }
fn = func() int { return 2 }
__check(fmt.Sprint(fn()), "2") }
