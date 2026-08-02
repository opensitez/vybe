// vybe-test: go/blank_identifier_extended/blank_discard_func_call_returns
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func divmod(a int, b int) (int, int) { return a / b, a % b }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { q, _ := divmod(10, 3)
__check(fmt.Sprint(q), "3") }
