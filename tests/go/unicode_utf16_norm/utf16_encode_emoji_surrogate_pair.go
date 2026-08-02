// vybe-test: go/unicode_utf16_norm/utf16_encode_emoji_surrogate_pair
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

func main() { u := utf16.Encode([]rune("🙂"))
__check(fmt.Sprint(len(u)), "2")
__check(fmt.Sprint(int(u[0])), "55357")
__check(fmt.Sprint(int(u[1])), "56898") }
