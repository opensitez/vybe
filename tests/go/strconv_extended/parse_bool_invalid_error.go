// vybe-test: go/strconv_extended/parse_bool_invalid_error
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

func main() { _, err := strconv.ParseBool("maybe")
__check(fmt.Sprint(err != nil), "true") }
