// vybe-test: go/constants_iota_advanced/iota_reset_in_new_const_block
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( A = iota; B = 10 )
const ( C = iota; D )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(B), "10")
__check(fmt.Sprint(C), "0")
__check(fmt.Sprint(D), "1") }
