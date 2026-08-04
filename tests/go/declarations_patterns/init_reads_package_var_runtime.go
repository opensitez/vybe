// vybe-test: go/declarations_patterns/init_reads_package_var_runtime
// origin: languages/go/tests/go/test_declarations_patterns.rs

package main
import "fmt"
var base = 6
var total int
func init() { total = base + 1 }
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

func main() { __p(fmt.Sprint(total))
__check("7")
}
