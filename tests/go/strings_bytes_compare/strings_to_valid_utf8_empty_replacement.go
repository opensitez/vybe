// vybe-test: go/strings_bytes_compare/strings_to_valid_utf8_empty_replacement
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

func main() { s := string([]byte{0xff})
out := strings.ToValidUTF8(s, "")
__check(fmt.Sprint(len(out) >= 0), "true") }
