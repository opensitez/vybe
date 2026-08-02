// vybe-test: go/strconv_extended/parse_bool_mixed_case_true
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

func main() { v, _ := strconv.ParseBool("True")
__check(fmt.Sprint(v), "true") }
