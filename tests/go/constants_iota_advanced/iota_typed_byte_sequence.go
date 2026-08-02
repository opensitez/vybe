// vybe-test: go/constants_iota_advanced/iota_typed_byte_sequence
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( First byte = iota; Second; Third )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(int(First)), "0")
__check(fmt.Sprint(int(Third)), "2") }
