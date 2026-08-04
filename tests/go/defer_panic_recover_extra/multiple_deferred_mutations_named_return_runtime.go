// vybe-test: go/defer_panic_recover_extra/multiple_deferred_mutations_named_return_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func build() (result int) { defer func() { result += 3 }()
defer func() { result *= 2 }()
result = 4
return }
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

func main() { __p(fmt.Sprint(build()))
__check("11")
}
