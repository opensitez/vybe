// vybe-test: go/strings_ops_extended/index_func_first_uppercase
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

func main() { __check(fmt.Sprint(strings.IndexFunc("vybeGo", func(r rune) bool { return r >= 'A' && r <= 'Z' })), "4") }
