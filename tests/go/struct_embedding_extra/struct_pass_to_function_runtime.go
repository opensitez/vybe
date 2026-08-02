// vybe-test: go/struct_embedding_extra/struct_pass_to_function_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type point struct { x int
y int }
func total(p point) int { return p.x + p.y }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(total(point{x: 2, y: 5})), "7")
}
