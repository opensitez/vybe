// vybe-test: go/constants_iota_advanced/iota_combined_with_const_outside
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const offset = 3
const ( A = iota + offset; B )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(A), "3")
__check(fmt.Sprint(B), "4") }
