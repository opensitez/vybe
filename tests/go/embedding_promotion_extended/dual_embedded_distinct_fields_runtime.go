// vybe-test: go/embedding_promotion_extended/dual_embedded_distinct_fields_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

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

func main() { p := point{axis: axis{x: 2}, ord: ord{y: 5}}
__check(fmt.Sprint(p.x + p.y), "7") }
