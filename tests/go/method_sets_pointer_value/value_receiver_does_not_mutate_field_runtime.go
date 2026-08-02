// vybe-test: go/method_sets_pointer_value/value_receiver_does_not_mutate_field_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type cell struct { n int }
func (c cell) bump() { c.n++ }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { v := cell{n: 5}
v.bump()
__check(fmt.Sprint(v.n), "5") }
