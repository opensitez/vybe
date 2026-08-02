// vybe-test: go/unicode_utf8/utf8_encode_rune_ascii
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

func main() { buf := make([]byte, 4)
n := utf8.EncodeRune(buf, 'A')
__check(fmt.Sprint(n), "1")
__check(fmt.Sprint(int(buf[0])), "65") }
