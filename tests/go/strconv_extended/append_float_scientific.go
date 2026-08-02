// vybe-test: go/strconv_extended/append_float_scientific
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

func main() { b := strconv.AppendFloat([]byte{}, 100.0, 'e', 0, 64)
__check(fmt.Sprint(len(string(b)) > 0), "true") }
