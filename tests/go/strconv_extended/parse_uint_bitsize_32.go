// vybe-test: go/strconv_extended/parse_uint_bitsize_32
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

func main() { n, _ := strconv.ParseUint("4294967295", 10, 32)
__check(fmt.Sprint(n), "4294967295") }
