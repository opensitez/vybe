// vybe-test: go/strings_ops_extended/replace_limited_count
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

func main() { __check(fmt.Sprint(strings.Replace("aabbcc", "b", "x", 1)), "aaxbcc") }
