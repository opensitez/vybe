// vybe-test: go/constants_iota_advanced/iota_byte_hex_shift
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( H0 byte = 0x10 << iota; H1; H2 )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(int(H0)), "16")
__check(fmt.Sprint(int(H2)), "64") }
