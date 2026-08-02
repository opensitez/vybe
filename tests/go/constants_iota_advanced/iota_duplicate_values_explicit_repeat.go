// vybe-test: go/constants_iota_advanced/iota_duplicate_values_explicit_repeat
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( Low = iota; Mid = Low; High = iota )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Low), "0")
__check(fmt.Sprint(Mid), "0")
__check(fmt.Sprint(High), "2") }
