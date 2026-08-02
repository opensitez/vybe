// vybe-test: go/function_literals_closures/closure_with_named_result_params
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { divide := func(a, b int) (q int, r int) { q = a / b
r = a % b
return }
q, r := divide(10, 3)
__check(fmt.Sprint(q), "3")
__check(fmt.Sprint(r), "1") }
