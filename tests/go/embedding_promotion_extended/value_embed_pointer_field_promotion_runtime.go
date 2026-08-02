// vybe-test: go/embedding_promotion_extended/value_embed_pointer_field_promotion_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type a struct { x int }
type b struct { *a }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { v := b{a: &a{x: 6}}
__check(fmt.Sprint(v.x), "6") }
