// vybe-test: go/constants_iota_advanced/iota_power_of_three
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( P0 = 1; P1 = 3 * iota; P2 )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(P0), "1")
__check(fmt.Sprint(P1), "3")
__check(fmt.Sprint(P2), "6") }
