// vybe-test: go/strconv_extended/parse_uint_bitsize_8_max
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

func main() { n, _ := strconv.ParseUint("255", 10, 8)
__check(fmt.Sprint(n), "255") }
