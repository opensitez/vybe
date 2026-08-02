// vybe-test: go/strings_ops_extended/last_index_func_trailing_digit
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

func main() { __check(fmt.Sprint(strings.LastIndexFunc("ab12", func(r rune) bool { return r >= '0' && r <= '9' })), "3") }
