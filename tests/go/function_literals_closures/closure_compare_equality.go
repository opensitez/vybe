// vybe-test: go/function_literals_closures/closure_compare_equality
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := func() {}
b := a
__check(fmt.Sprint(a == nil), "false")
__check(fmt.Sprint(b == nil), "false") }
