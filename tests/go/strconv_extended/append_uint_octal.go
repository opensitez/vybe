// vybe-test: go/strconv_extended/append_uint_octal
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

func main() { b := strconv.AppendUint([]byte{}, 8, 8)
__check(fmt.Sprint(string(b)), "10") }
