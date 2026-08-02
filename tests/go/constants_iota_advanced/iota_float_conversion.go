// vybe-test: go/constants_iota_advanced/iota_float_conversion
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( F0 = float64(iota); F1; F2 )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(F0), "0")
__check(fmt.Sprint(F2), "2") }
