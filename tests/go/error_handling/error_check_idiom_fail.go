// vybe-test: go/error_handling/error_check_idiom_fail
// origin: languages/go/tests/go/test_error_handling.rs

package main
import "fmt"
type BasicErr struct{}
func (BasicErr) Error() string { return "err" }
func divide(a int, b int) (int, error) { if b == 0 { return 0, BasicErr{} }
return a / b, nil }
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

func main() { res, err := divide(10, 0)
if err != nil { __p(fmt.Sprint("error")) } else { __p(fmt.Sprint(res)) } 
__check("error")
}
