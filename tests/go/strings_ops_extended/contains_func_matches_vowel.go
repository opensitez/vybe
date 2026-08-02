// vybe-test: go/strings_ops_extended/contains_func_matches_vowel
// origin: languages/go/tests/go/test_strings_ops_extended.rs

package main
import "fmt"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(strings.ContainsFunc("rhythm", func(r rune) bool { return r == 'a' || r == 'e' || r == 'i' || r == 'o' || r == 'u' })), "false") }
