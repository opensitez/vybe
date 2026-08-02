// vybe-test: go/strings_bytes_compare/strings_map_unicode_lower
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

func main() { out := strings.Map(func(r rune) rune { if r >= 'A' && r <= 'Z' { return r + 32 }; return r }, "GoLang")
__check(fmt.Sprint(out), "golang") }
