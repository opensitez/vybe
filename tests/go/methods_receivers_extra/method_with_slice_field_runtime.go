// vybe-test: go/methods_receivers_extra/method_with_slice_field_runtime
// origin: languages/go/tests/go/test_methods_receivers_extra.rs

package main
import "fmt"
type bag struct { values []int }
func (b bag) count() int { return len(b.values) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := bag{values: []int{1, 2, 3}}
__check(fmt.Sprint(value.count()), "3")
}
