// vybe-test: go/struct_embedding_extra/struct_swap_values_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type point struct { x int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { left := point{x: 1}
right := point{x: 9}
left, right = right, left
__check(fmt.Sprint(left.x), "9")
__check(fmt.Sprint(right.x), "1")
}
