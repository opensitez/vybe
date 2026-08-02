// vybe-test: go/strings_bytes_compare/strings_map_identity
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

func main() { out := strings.Map(func(r rune) rune { return r }, "abc")
__check(fmt.Sprint(out), "abc") }
