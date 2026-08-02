// vybe-test: go/strings_bytes_compare/strings_to_valid_utf8_keeps_valid
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

func main() { out := strings.ToValidUTF8("go", "?")
__check(fmt.Sprint(out), "go") }
