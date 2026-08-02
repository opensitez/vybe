// vybe-test: go/constants_iota_advanced/iota_blank_then_explicit
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( _ = iota; Z = iota + 5 )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Z), "6") }
