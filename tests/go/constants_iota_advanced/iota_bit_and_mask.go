// vybe-test: go/constants_iota_advanced/iota_bit_and_mask
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( M0 = 1 << iota; M1; M2 )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(M0 & M1), "0")
__check(fmt.Sprint(M0 | M1), "3") }
