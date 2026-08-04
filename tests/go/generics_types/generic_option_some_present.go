// vybe-test: go/generics_types/generic_option_some_present
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Option[T any] struct { Value T
Present bool }
func Some[T any](v T) Option[T] { return Option[T]{Value: v, Present: true} }
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

func main() { o := Some(42)
__p(fmt.Sprint(o.Present))
__p(fmt.Sprint(o.Value)) 
__check("true\n42")
}
