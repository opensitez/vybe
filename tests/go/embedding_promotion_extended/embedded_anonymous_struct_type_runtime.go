// vybe-test: go/embedding_promotion_extended/embedded_anonymous_struct_type_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type outer struct { struct { x int
y int } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { o := outer{struct { x int
y int }{x: 1, y: 2}}
__check(fmt.Sprint(o.x + o.y), "3") }
