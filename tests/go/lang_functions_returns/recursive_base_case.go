// vybe-test: go/lang_functions_returns/recursive_base_case
// origin: languages/go/tests/go/test_lang_functions_returns.rs

package main
import "fmt"
func fact(n int) int { if n <= 1 { return 1 }
return n * fact(n-1) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(fact(4)), "24") }
