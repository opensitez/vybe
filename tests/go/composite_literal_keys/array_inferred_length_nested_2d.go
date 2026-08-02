// vybe-test: go/composite_literal_keys/array_inferred_length_nested_2d
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { grid := [...][2]int{{0: 1, 1: 2}, {1: 4, 0: 3}}
__check(fmt.Sprint(grid[1][0]), "3")
__check(fmt.Sprint(grid[0][1]), "2")
}
