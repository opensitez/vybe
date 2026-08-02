// vybe-test: go/constants_iota_advanced/iota_bit_flags_or_combined
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( Read = 1 << iota; Write; Execute )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Read | Write | Execute), "7") }
