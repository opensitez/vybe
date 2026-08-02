// vybe-test: go/pointer_receivers_advanced/new_struct_zero_value_matches_literal_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type widget struct { size int
label string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { fromNew := new(widget)
fromLit := &widget{}
__check(fmt.Sprint(fromNew.size == fromLit.size), "true")
__check(fmt.Sprint(fromNew.label == fromLit.label), "true")
}
