// vybe-test: go/constants_iota_advanced/iota_xor_toggle_bits
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( B0 = 1 << iota; B1 )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(B0 ^ B1), "3") }
