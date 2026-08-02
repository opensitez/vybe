// vybe-test: go/pointer_receivers_advanced/value_receiver_slice_append_not_visible_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type bag struct { items []int }
func (b bag) push(v int) { b.items = append(b.items, v) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := bag{items: []int{1}}
value.push(2)
__check(fmt.Sprint(len(value.items)), "1")
}
