// vybe-test: go/unicode_utf8/utf8_decode_rune_ascii
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

func main() { r, size := utf8.DecodeRune([]byte("Z"))
__check(fmt.Sprint(int(r)), "90")
__check(fmt.Sprint(size), "1") }
