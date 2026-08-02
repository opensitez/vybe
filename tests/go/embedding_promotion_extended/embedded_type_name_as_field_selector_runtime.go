// vybe-test: go/embedding_promotion_extended/embedded_type_name_as_field_selector_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type coords struct { x int
y int }
type point struct { coords }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { p := point{coords: coords{x: 3, y: 4}}
__check(fmt.Sprint(p.coords.x), "3") }
