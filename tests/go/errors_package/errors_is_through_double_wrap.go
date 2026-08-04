// vybe-test: go/errors_package/errors_is_through_double_wrap
// origin: languages/go/tests/go/test_errors_package.rs

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

func main() { err := fmt.Errorf("layer1: %w", fmt.Errorf("layer2: %w", ErrBase))
__p(fmt.Sprint(errors.Is(err, ErrBase))) 
__check("true")
}
