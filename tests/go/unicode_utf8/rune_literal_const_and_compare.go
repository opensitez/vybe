// vybe-test: go/unicode_utf8/rune_literal_const_and_compare
// origin: languages/go/tests/go/test_unicode_utf8.rs

package main
import "fmt"
const letter rune = 'λ'
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(int(letter)), "955")
__check(fmt.Sprint(letter == '\u03BB'), "true") }
