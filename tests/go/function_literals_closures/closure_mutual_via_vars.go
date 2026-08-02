// vybe-test: go/function_literals_closures/closure_mutual_via_vars
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var even func(int) bool
var odd func(int) bool
even = func(n int) bool { if n == 0 { return true }
return odd(n-1) }
odd = func(n int) bool { if n == 0 { return false }
return even(n-1) }
__check(fmt.Sprint(even(4)), "true")
__check(fmt.Sprint(odd(3)), "true") }
