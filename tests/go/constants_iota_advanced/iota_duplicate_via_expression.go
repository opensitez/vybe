// vybe-test: go/constants_iota_advanced/iota_duplicate_via_expression
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( X = iota; Y = X + 0; Z = iota )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(X), "0")
__check(fmt.Sprint(Y), "0")
__check(fmt.Sprint(Z), "2") }
