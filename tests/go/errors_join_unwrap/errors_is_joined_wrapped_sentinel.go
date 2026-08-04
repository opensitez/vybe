// vybe-test: go/errors_join_unwrap/errors_is_joined_wrapped_sentinel
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
var ErrBase = errors.New("base")
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

func main() { wrapped := fmt.Errorf("w: %w", ErrBase)
joined := errors.Join(wrapped, errors.New("other"))
__p(fmt.Sprint(errors.Is(joined, ErrBase))) 
__check("true")
}
