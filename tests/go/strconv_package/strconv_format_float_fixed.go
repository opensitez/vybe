// vybe-test: go/strconv_package/strconv_format_float_fixed
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

func main() { __check(fmt.Sprint(strconv.FormatFloat(1.5, 'f', 1, 64)), "1.5") }
