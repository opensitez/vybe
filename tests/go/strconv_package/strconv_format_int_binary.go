// vybe-test: go/strconv_package/strconv_format_int_binary
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

func main() { __check(fmt.Sprint(strconv.FormatInt(5, 2)), "101") }
