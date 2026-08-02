// vybe-test: go/constants_iota_advanced/iota_offset_add_ten
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( Base = iota + 10; Next; Last )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Base), "10")
__check(fmt.Sprint(Last), "12") }
