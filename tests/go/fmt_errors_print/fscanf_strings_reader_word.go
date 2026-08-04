// vybe-test: go/fmt_errors_print/fscanf_strings_reader_word
// origin: languages/go/tests/go/test_fmt_errors_print.rs

package main
import "fmt"
import "strings"
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { var s string
c, _ := fmt.Fscanf(strings.NewReader("vybe"), "%s", &s)
__p(fmt.Sprint(c) + " " + fmt.Sprint(s)) 
__check("1 vybe")
}
