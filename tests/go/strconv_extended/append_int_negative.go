// vybe-test: go/strconv_extended/append_int_negative
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

func main() { b := strconv.AppendInt([]byte{}, -8, 10)
__check(fmt.Sprint(string(b)), "-8") }
