// vybe-test: go/strconv_extended/format_float_hex_prefix
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

func main() { s := strconv.FormatFloat(10.0, 'x', 0, 64)
__check(fmt.Sprint(len(s) > 2), "true") }
