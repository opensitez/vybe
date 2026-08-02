// vybe-test: go/fmt_errors_print/sscanf_single_string
// origin: languages/go/tests/go/test_fmt_errors_print.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var s string
c, _ := fmt.Sscanf("hello", "%s", &s)
__check(fmt.Sprint(c) + " " + fmt.Sprint(s), "1 hello") }
