// vybe-test: go/unicode_utf16_norm/utf8_encode_rune_two_byte
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

func main() { buf := make([]byte, 4)
n := utf8.EncodeRune(buf, 'é')
__check(fmt.Sprint(n), "2")
__check(fmt.Sprint(int(buf[0])), "195") }
