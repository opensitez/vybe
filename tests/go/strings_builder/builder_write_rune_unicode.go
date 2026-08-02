// vybe-test: go/strings_builder/builder_write_rune_unicode
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

func main() { var b strings.Builder
b.WriteRune('日')
__check(fmt.Sprint(b.String()), "日") }
