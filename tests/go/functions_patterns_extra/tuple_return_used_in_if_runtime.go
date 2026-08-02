// vybe-test: go/functions_patterns_extra/tuple_return_used_in_if_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func dims() (int, int) { return 2, 4 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { w, h := dims()
if w < h { __check(fmt.Sprint(h - w), "2") } }
