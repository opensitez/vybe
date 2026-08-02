// vybe-test: go/strconv_extended/parse_uint_bitsize_16
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

func main() { n, _ := strconv.ParseUint("65535", 10, 16)
__check(fmt.Sprint(n), "65535") }
