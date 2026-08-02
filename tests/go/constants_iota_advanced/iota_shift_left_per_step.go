// vybe-test: go/constants_iota_advanced/iota_shift_left_per_step
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( Bit0 = 1 << iota; Bit1; Bit2; Bit3 )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Bit0), "1")
__check(fmt.Sprint(Bit3), "8") }
