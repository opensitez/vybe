// vybe-test: go/methods_receivers_extra/method_returns_struct_field_runtime
// origin: languages/go/tests/go/test_methods_receivers_extra.rs

package main
import "fmt"
type point struct { x int }
func (p point) value() int { return p.x }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(point{x: 12}.value()), "12")
}
