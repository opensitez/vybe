// vybe-test: go/unicode_utf16_norm/utf16_decode_replaces_invalid_surrogate
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

func main() { rs := utf16.Decode([]uint16{0xD800})
__check(fmt.Sprint(int(rs[0])), "65533") }
