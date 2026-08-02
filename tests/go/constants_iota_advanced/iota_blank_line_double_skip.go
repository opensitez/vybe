// vybe-test: go/constants_iota_advanced/iota_blank_line_double_skip
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( _ = iota; _; X; Y )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(X), "2")
__check(fmt.Sprint(Y), "3") }
