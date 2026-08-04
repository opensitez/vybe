// vybe-test: go/method_sets_pointer_value/interface_assign_from_address_of_local_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type saver interface { save(int) }
type disk struct { used int }
func (d *disk) save(v int) { d.used += v }
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

func main() { local := disk{}
var s saver = &local
s.save(6)
__p(fmt.Sprint(local.used)) 
__check("6")
}
