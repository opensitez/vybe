// vybe-test: go/strconv_package/strconv_parse_bool_false
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

func main() { v, _ := strconv.ParseBool("false")
__check(fmt.Sprint(v), "false") }
