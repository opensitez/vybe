// vybe-test: go/defer_panic_recover_extra/defer_mutate_pointer_param_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
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

func main() { value := 1
func() { defer func(ptr *int) { *ptr = 5 }(&value) }()
__p(fmt.Sprint(value))
__check("5")
}
