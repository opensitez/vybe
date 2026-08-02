// vybe-test: go/constants_iota_advanced/iota_arithmetic_difference
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( A = 10 - iota; B; C )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(A), "10")
__check(fmt.Sprint(C), "8") }
