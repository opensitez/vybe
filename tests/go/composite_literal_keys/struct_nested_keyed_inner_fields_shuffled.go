// vybe-test: go/composite_literal_keys/struct_nested_keyed_inner_fields_shuffled
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type coord struct { x int
y int }
type rect struct { origin coord
size coord }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { r := rect{size: coord{y: 4, x: 3}, origin: coord{x: 1, y: 2}}
__check(fmt.Sprint(r.origin.x), "1")
__check(fmt.Sprint(r.size.y), "4")
}
