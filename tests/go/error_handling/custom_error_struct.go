// vybe-test: go/error_handling/custom_error_struct
// origin: languages/go/tests/go/test_error_handling.rs

package main
import "fmt"
type MyErr struct { msg string }
func (e MyErr) Error() string { return e.msg }
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

func main() { var err error = MyErr{msg: "failed"}
__p(fmt.Sprint(err.Error()))
__check("failed")
}
