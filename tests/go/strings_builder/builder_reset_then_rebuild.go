// vybe-test: go/strings_builder/builder_reset_then_rebuild
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
b.WriteString("old")
b.Reset()
b.WriteString("new")
__check(fmt.Sprint(b.String()), "new") }
