// vybe-test: go/iota_enumerations/iota_simple_three_values
// origin: languages/go/tests/go/test_iota_enumerations.rs

package main
import "fmt"
const ( A = iota; B; C )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(A), "0")
__check(fmt.Sprint(B), "1")
__check(fmt.Sprint(C), "2") }
