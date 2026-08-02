// vybe-test: go/strings_builder/replacer_write_string_count
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

func main() { rep := strings.NewReplacer("x", "y")
var b strings.Builder
n, _ := rep.WriteString(&b, "x")
__check(fmt.Sprint(n), "1")
__check(fmt.Sprint(b.String()), "y") }
