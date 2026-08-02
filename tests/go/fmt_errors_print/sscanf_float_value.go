// vybe-test: go/fmt_errors_print/sscanf_float_value
// origin: languages/go/tests/go/test_fmt_errors_print.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var f float64
c, _ := fmt.Sscanf("3.14", "%f", &f)
__check(fmt.Sprint(c) + " " + fmt.Sprint(f), "1 3.14") }
