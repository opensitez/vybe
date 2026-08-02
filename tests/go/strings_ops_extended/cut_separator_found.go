// vybe-test: go/strings_ops_extended/cut_separator_found
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

func main() { before, after, found := strings.Cut("hello,world", ",")
__check(fmt.Sprint(before), "hello")
__check(fmt.Sprint(after), "world")
__check(fmt.Sprint(found), "true") }
