// vybe-test: go/composite_literal_keys/array_of_structs_mixed_keyed_unkeyed_elements
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type pair struct { left int
right int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := [3]pair{{right: 2, left: 1}, {3, 4}, pair{left: 5, right: 6}}
__check(fmt.Sprint(a[0].left), "1")
__check(fmt.Sprint(a[1].right), "4")
__check(fmt.Sprint(a[2].left), "5")
}
