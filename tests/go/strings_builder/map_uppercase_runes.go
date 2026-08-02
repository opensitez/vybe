// vybe-test: go/strings_builder/map_uppercase_runes
// origin: languages/go/tests/go/test_strings_builder.rs

package main
import "fmt"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { out := strings.Map(func(r rune) rune { if r >= 'a' && r <= 'z' { return r - 32 }; return r }, "AbC")
__check(fmt.Sprint(out), "ABC") }
