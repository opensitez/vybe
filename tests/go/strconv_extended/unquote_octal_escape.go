// vybe-test: go/strconv_extended/unquote_octal_escape
// origin: languages/go/tests/go/test_strconv_extended.rs

package main
import "fmt"
import "strconv"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s, _ := strconv.Unquote(`"\101"`)
__check(fmt.Sprint(s), "A") }
