// vybe-test: go/unicode_utf16_norm/utf8_decode_rune_in_string
// origin: languages/go/tests/go/test_unicode_utf16_norm.rs

package main
import "fmt"
import "unicode/utf8"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { r, size := utf8.DecodeRuneInString("世go")
__check(fmt.Sprint(int(r)), "19990")
__check(fmt.Sprint(size), "3") }
