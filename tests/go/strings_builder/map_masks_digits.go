// vybe-test: go/strings_builder/map_masks_digits
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

func main() { out := strings.Map(func(r rune) rune { if r >= '0' && r <= '9' { return '#' }; return r }, "a1b2")
__check(fmt.Sprint(out), "a#b#") }
