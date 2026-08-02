// vybe-test: go/iota_enumerations/iota_skip_with_blank
// origin: languages/go/tests/go/test_iota_enumerations.rs

package main
import "fmt"
const ( _ = iota; X = iota; Y )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(X), "1")
__check(fmt.Sprint(Y), "2") }
