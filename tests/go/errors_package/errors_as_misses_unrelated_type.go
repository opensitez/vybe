// vybe-test: go/errors_package/errors_as_misses_unrelated_type
// origin: languages/go/tests/go/test_errors_package.rs

package main
import "fmt"
import "errors"
type coded struct { n int }
func (c coded) Error() string { return "coded" }
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

func main() { err := errors.New("plain")
var target coded
__p(fmt.Sprint(errors.As(err, &target))) 
__check("false")
}
