// vybe-test: go/strings_ops_extended/trim_prefix_removed
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

func main() { rest := strings.TrimPrefix("prefix:value", "prefix:")
__check(fmt.Sprint(rest), "value") }
