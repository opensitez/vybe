// vybe-test: go/unicode_package/unicode_is_letter_ascii_upper
// origin: languages/go/tests/go/test_unicode_package.rs

package main
import "fmt"
import "unicode"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(unicode.IsLetter('A')), "true") }
