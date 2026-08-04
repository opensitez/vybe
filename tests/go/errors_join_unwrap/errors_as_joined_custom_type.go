// vybe-test: go/errors_join_unwrap/errors_as_joined_custom_type
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
type coded struct { code int }
func (c coded) Error() string { return fmt.Sprintf("code %d", c.code) }
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

func main() { inner := coded{code: 42}
joined := errors.Join(inner, errors.New("plain"))
var target coded
__p(fmt.Sprint(errors.As(joined, &target)))
__p(fmt.Sprint(target.code)) 
__check("true\n42")
}
