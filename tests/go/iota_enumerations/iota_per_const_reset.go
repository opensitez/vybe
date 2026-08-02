// vybe-test: go/iota_enumerations/iota_per_const_reset
// origin: languages/go/tests/go/test_iota_enumerations.rs

package main
import "fmt"
const ( A = iota; B )
const ( C = iota; D )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(B), "1")
__check(fmt.Sprint(D), "1") }
