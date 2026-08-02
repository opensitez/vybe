// vybe-test: go/strings_builder/builder_write_byte
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
b.WriteByte('A')
__check(fmt.Sprint(b.String()), "A") }
