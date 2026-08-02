// vybe-test: go/method_sets_pointer_value/nested_embedded_value_method_two_levels_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type leaf struct { v int }
func (l leaf) val() int { return l.v }
type branch struct { leaf }
type trunk struct { branch }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { t := trunk{branch: branch{leaf: leaf{v: 9}}}
__check(fmt.Sprint(t.val()), "9") }
