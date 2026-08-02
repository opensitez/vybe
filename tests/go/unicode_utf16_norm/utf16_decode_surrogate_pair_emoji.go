// vybe-test: go/unicode_utf16_norm/utf16_decode_surrogate_pair_emoji
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

func main() { rs := utf16.Decode([]uint16{0xD83D, 0xDE42})
__check(fmt.Sprint(len(rs)), "1")
__check(fmt.Sprint(int(rs[0])), "128578") }
