// vybe-test: go/unicode_utf8/utf8_decode_rune_invalid_byte
// origin: languages/go/tests/go/test_unicode_utf8.rs

package main
import "fmt"
import "unicode/utf8"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { r, size := utf8.DecodeRune([]byte{0xff})
__check(fmt.Sprint(int(r)), "65533")
__check(fmt.Sprint(size), "1") }
