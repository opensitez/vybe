// vybe-test: go/strings_ops_extended/cut_prefix_no_match_returns_original
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

func main() { rest, found := strings.CutPrefix("ftp://host", "https://")
__check(fmt.Sprint(rest), "ftp://host")
__check(fmt.Sprint(found), "false") }
