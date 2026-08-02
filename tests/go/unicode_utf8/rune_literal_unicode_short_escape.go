// vybe-test: go/unicode_utf8/rune_literal_unicode_short_escape
// origin: languages/go/tests/go/test_unicode_utf8.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(int('\u03BB')), "955") }
