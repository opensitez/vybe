// vybe-test: go/strconv_extended/parse_float_inf_negative
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

func main() { v, _ := strconv.ParseFloat("-Inf", 64)
__check(fmt.Sprint(v < -1e308), "true") }
