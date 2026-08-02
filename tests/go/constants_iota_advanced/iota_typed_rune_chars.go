// vybe-test: go/constants_iota_advanced/iota_typed_rune_chars
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( Alpha rune = 'A' + iota; Beta; Gamma )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(int(Alpha)), "65")
__check(fmt.Sprint(int(Gamma)), "67") }
