// vybe-test: go/strconv_package/strconv_parse_int_base10
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

func main() { n, _ := strconv.ParseInt("42", 10, 64)
__check(fmt.Sprint(n), "42") }
