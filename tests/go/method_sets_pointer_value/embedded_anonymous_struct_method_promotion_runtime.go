// vybe-test: go/method_sets_pointer_value/embedded_anonymous_struct_method_promotion_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type coords struct { x int
y int }
func (c coords) sum() int { return c.x + c.y }
type point struct { coords }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { p := point{coords: coords{x: 2, y: 5}}
__check(fmt.Sprint(p.sum()), "7") }
