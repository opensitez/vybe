// vybe-test: go/fmt_errors_print/fscanf_strings_reader_word
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

func main() { var s string
c, _ := fmt.Fscanf(strings.NewReader("vybe"), "%s", &s)
__check(fmt.Sprint(c) + " " + fmt.Sprint(s), "1 vybe") }
