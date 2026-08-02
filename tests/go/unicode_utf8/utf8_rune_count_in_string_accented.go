// vybe-test: go/unicode_utf8/utf8_rune_count_in_string_accented
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

func main() { __check(fmt.Sprint(utf8.RuneCountInString("héllo")), "5") }
