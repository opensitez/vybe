// vybe-test: go/method_sets_pointer_value/pointer_only_interface_from_local_address_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type mutator interface { mutate() }
type data struct { n int }
func (d *data) mutate() { d.n = 99 }
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

func main() { local := data{n: 1}
var m mutator = &local
m.mutate()
__p(fmt.Sprint(local.n)) 
__check("99")
}
