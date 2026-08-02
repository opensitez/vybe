// vybe-test: go/constants_iota_advanced/iota_string_number_interleave
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( Name = "v"; Code = iota; Next )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Name), "v")
__check(fmt.Sprint(Code), "0")
__check(fmt.Sprint(Next), "1") }
