// vybe-test: go/fmt_errors_print/fscanf_int_and_string_pair
// origin: languages/go/tests/go/test_fmt_errors_print.rs

package main
import "fmt"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var n int
var s string
c, _ := fmt.Fscanf(strings.NewReader("3 ok"), "%d %s", &n, &s)
__check(fmt.Sprint(c) + " " + fmt.Sprint(n) + " " + fmt.Sprint(s), "2 3 ok") }
