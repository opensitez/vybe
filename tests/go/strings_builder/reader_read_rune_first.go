// vybe-test: go/strings_builder/reader_read_rune_first
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

func main() { r := strings.NewReader("日")
ch, _, _ := r.ReadRune()
__check(fmt.Sprint(string(ch)), "日") }
