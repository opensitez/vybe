// vybe-test: go/constants_iota_advanced/iota_parenthesized_expression
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( V = (iota + 1) * 2; W; X )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(V), "2")
__check(fmt.Sprint(X), "6") }
