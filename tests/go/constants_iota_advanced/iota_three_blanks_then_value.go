// vybe-test: go/constants_iota_advanced/iota_three_blanks_then_value
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( _ = iota; _; _; Target )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Target), "3") }
