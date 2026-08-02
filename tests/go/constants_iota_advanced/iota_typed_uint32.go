// vybe-test: go/constants_iota_advanced/iota_typed_uint32
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( U0 uint32 = iota; U1; U2 )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(U0), "0")
__check(fmt.Sprint(U2), "2") }
