// vybe-test: go/constants_iota_advanced/iota_mixed_expression_and_implicit
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( Seed = iota * 2; Step; Tail )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Seed), "0")
__check(fmt.Sprint(Tail), "4") }
