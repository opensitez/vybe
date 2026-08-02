// vybe-test: go/strconv_extended/parse_int_bitsize_16
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

func main() { n, _ := strconv.ParseInt("-32768", 10, 16)
__check(fmt.Sprint(n), "-32768") }
