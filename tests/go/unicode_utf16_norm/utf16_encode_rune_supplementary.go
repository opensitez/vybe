// vybe-test: go/unicode_utf16_norm/utf16_encode_rune_supplementary
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

func main() { u1, u2 := utf16.EncodeRune(0x1F600)
__check(fmt.Sprint(int(u1)), "55357")
__check(fmt.Sprint(int(u2)), "56832") }
