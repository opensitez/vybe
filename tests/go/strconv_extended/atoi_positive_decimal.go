// vybe-test: go/strconv_extended/atoi_positive_decimal
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

func main() { n, err := strconv.Atoi("12345")
__check(fmt.Sprint(n), "12345")
__check(fmt.Sprint(err == nil), "true") }
