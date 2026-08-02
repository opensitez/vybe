// vybe-test: go/fmt_errors_print/sscanf_bool_true
// origin: languages/go/tests/go/test_fmt_errors_print.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var ok bool
c, _ := fmt.Sscanf("true", "%t", &ok)
__check(fmt.Sprint(c) + " " + fmt.Sprint(ok), "1 true") }
