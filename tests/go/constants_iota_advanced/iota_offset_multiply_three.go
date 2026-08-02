// vybe-test: go/constants_iota_advanced/iota_offset_multiply_three
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( Start = iota * 3; Mid; End )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Start), "0")
__check(fmt.Sprint(End), "6") }
