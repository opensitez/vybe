// vybe-test: go/init_blank_import/init_order_three_sequential_appends
// origin: languages/go/tests/go/test_init_blank_import.rs

package main
import "fmt"
var order string
func init() { order = order + "1" }
func init() { order = order + "2" }
func init() { order = order + "3" }
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

func main() { __p(fmt.Sprint(order)) 
__check("123")
}
