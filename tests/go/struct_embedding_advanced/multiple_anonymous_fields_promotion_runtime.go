// vybe-test: go/struct_embedding_advanced/multiple_anonymous_fields_promotion_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

package main
import "fmt"
type axis struct { x int }
type ord struct { y int }
type point struct { axis
ord }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := point{axis: axis{x: 4}, ord: ord{y: 6}}
__check(fmt.Sprint(value.x + value.y), "10")
}
