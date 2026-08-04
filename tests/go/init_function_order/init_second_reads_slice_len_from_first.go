// vybe-test: go/init_function_order/init_second_reads_slice_len_from_first
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var items []string
var size int
func init() { items = []string{"a", "b"} }
func init() { size = len(items) }
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

func main() { __p(fmt.Sprint(size)) 
__check("2")
}
