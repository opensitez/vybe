// vybe-test: go/constants_iota_advanced/iota_rune_offset_from_a
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( R0 rune = 'a' + iota; R1; R2 )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(string(R0)), "a")
__check(fmt.Sprint(string(R2)), "c") }
