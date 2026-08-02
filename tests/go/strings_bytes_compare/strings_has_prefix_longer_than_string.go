// vybe-test: go/strings_bytes_compare/strings_has_prefix_longer_than_string
// origin: languages/go/tests/go/test_strings_bytes_compare.rs

package main
import "fmt"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(strings.HasPrefix("go", "gopher")), "false") }
