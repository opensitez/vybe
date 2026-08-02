// vybe-test: go/unicode_utf16_norm/utf16_encode_rune_ascii
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

func main() { u1, u2 := utf16.EncodeRune('Z')
__check(fmt.Sprint(u1), "90")
__check(fmt.Sprint(u2), "65535") }
