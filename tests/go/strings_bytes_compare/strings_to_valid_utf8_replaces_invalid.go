// vybe-test: go/strings_bytes_compare/strings_to_valid_utf8_replaces_invalid
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

func main() { s := string([]byte{0xff, 0xfe, 'a'})
out := strings.ToValidUTF8(s, "?")
__check(fmt.Sprint(len(out) > 1), "true") }
