// vybe-test: go/unicode_utf8/utf8_encode_rune_three_byte
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
n := utf8.EncodeRune(buf, '世')
__check(fmt.Sprint(n), "3")
__check(fmt.Sprint(int(buf[0])), "228")
__check(fmt.Sprint(int(buf[1])), "184")
__check(fmt.Sprint(int(buf[2])), "150") }
