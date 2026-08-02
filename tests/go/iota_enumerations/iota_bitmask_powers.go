// vybe-test: go/iota_enumerations/iota_bitmask_powers
// origin: languages/go/tests/go/test_iota_enumerations.rs

package main
import "fmt"
const ( FlagA = 1 << iota; FlagB; FlagC )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(FlagA), "1")
__check(fmt.Sprint(FlagB), "2")
__check(fmt.Sprint(FlagC), "4") }
