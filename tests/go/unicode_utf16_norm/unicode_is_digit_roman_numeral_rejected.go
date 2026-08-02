// vybe-test: go/unicode_utf16_norm/unicode_is_digit_roman_numeral_rejected
// origin: languages/go/tests/go/test_unicode_utf16_norm.rs

package main
import "fmt"
import "unicode"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(unicode.IsDigit('Ⅳ')), "false") }
