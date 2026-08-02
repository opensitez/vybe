// vybe-test: go/unicode_utf16_norm/utf16_decode_rune_single_unit
// origin: languages/go/tests/go/test_unicode_utf16_norm.rs

package main
import "fmt"
import "unicode/utf16"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { r := utf16.DecodeRune(65, 65535)
__check(fmt.Sprint(int(r)), "65") }
