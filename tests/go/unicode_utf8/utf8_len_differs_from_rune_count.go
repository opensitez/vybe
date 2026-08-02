// vybe-test: go/unicode_utf8/utf8_len_differs_from_rune_count
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

func main() { s := "日本"
__check(fmt.Sprint(len(s)), "6")
__check(fmt.Sprint(utf8.RuneCountInString(s)), "2") }
