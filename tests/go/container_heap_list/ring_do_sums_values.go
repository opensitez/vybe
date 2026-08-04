// vybe-test: go/container_heap_list/ring_do_sums_values
// origin: languages/go/tests/go/test_container_heap_list.rs

package main
import "fmt"
import "container/ring"
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

func main() { r := ring.New(4)
sum := 0
for i := 0; i < 4; i++ { r.Value = i + 1
r = r.Next() }
r.Do(func(v interface{}) { sum += v.(int) })
__p(fmt.Sprint(sum)) 
__check("10")
}
