// vybe-test: go/strings_ops_extended/cut_separator_missing
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

func main() { before, after, found := strings.Cut("gopher", ",")
__check(fmt.Sprint(before), "gopher")
__check(fmt.Sprint(after), "")
__check(fmt.Sprint(found), "false") }
