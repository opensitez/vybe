// vybe-test: go/slice_aliasing_extra/slice_of_structs_runtime
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs

package main
import "fmt"
type point struct { x int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := []point{{x: 8}, {x: 9}}
__check(fmt.Sprint(values[1].x), "9")
}
