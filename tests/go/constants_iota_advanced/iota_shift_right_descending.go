// vybe-test: go/constants_iota_advanced/iota_shift_right_descending
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( A = 8 >> iota; B; C )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(A), "8")
__check(fmt.Sprint(C), "2") }
