// vybe-test: go/strconv_package/strconv_unquote_quoted
// origin: languages/go/tests/go/test_strconv_package.rs

package main
import "fmt"
import "strconv"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s, _ := strconv.Unquote(`"go"`)
__check(fmt.Sprint(s), "go") }
