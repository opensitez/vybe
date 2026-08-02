// vybe-test: go/strings_bytes_compare/strings_count_single_char
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

func main() { __check(fmt.Sprint(strings.Count("banana", "a")), "3") }
