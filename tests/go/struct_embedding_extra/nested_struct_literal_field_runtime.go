// vybe-test: go/struct_embedding_extra/nested_struct_literal_field_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type point struct { x int
y int }
type box struct { p point }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := box{p: point{x: 3, y: 4}}
__check(fmt.Sprint(value.p.x + value.p.y), "7")
}
