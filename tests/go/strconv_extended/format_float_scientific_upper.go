// vybe-test: go/strconv_extended/format_float_scientific_upper
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

func main() { __check(fmt.Sprint(strconv.FormatFloat(0.0012, 'E', 1, 64)), "1.2E-03") }
