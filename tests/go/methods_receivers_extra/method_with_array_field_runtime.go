// vybe-test: go/methods_receivers_extra/method_with_array_field_runtime
// origin: languages/go/tests/go/test_methods_receivers_extra.rs

package main
import "fmt"
type bag struct { values [2]int }
func (b bag) second() int { return b.values[1] }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := bag{values: [2]int{3, 9}}
__check(fmt.Sprint(value.second()), "9")
}
