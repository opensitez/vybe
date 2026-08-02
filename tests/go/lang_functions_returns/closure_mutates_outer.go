// vybe-test: go/lang_functions_returns/closure_mutates_outer
// origin: languages/go/tests/go/test_lang_functions_returns.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { n := 0
f := func() { n++ }
f()
f()
__check(fmt.Sprint(n), "2") }
