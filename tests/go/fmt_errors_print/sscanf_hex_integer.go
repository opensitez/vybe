// vybe-test: go/fmt_errors_print/sscanf_hex_integer
// origin: languages/go/tests/go/test_fmt_errors_print.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var n int
c, _ := fmt.Sscanf("ff", "%x", &n)
__check(fmt.Sprint(c) + " " + fmt.Sprint(n), "1 255") }
