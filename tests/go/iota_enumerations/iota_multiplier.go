// vybe-test: go/iota_enumerations/iota_multiplier
// origin: languages/go/tests/go/test_iota_enumerations.rs

package main
import "fmt"
const ( KB = 1 << (10 * iota); MB )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(KB), "1024")
__check(fmt.Sprint(MB), "1048576") }
