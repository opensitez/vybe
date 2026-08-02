// vybe-test: go/method_sets_pointer_value/value_receiver_slice_header_no_mutation_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type bag struct { items []int }
func (b bag) appendItem(v int) { b.items = append(b.items, v) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b := bag{items: []int{1}}
b.appendItem(2)
__check(fmt.Sprint(len(b.items)), "1") }
