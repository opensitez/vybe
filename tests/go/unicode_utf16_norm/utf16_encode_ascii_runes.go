// vybe-test: go/unicode_utf16_norm/utf16_encode_ascii_runes
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

func main() { u := utf16.Encode([]rune("AB"))
__check(fmt.Sprint(len(u)), "2")
__check(fmt.Sprint(int(u[0])), "65")
__check(fmt.Sprint(int(u[1])), "66") }
