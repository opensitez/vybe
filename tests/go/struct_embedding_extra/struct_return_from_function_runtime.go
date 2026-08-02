// vybe-test: go/struct_embedding_extra/struct_return_from_function_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type point struct { x int
y int }
func build() point { return point{x: 4, y: 6} }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := build()
__check(fmt.Sprint(value.x), "4")
__check(fmt.Sprint(value.y), "6")
}
