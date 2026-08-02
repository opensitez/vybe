// vybe-test: go/fmt_errors_print/sscanf_int_then_string
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
var s string
c, _ := fmt.Sscanf("7 go", "%d %s", &n, &s)
__check(fmt.Sprint(c) + " " + fmt.Sprint(n) + " " + fmt.Sprint(s), "2 7 go") }
