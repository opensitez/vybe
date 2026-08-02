// vybe-test: go/embedding_promotion_extended/nested_value_embed_each_level_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type c struct { tag string }
type b struct { c }
type a struct { b }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { v := a{b: b{c: c{tag: "deep"}}}
__check(fmt.Sprint(v.tag), "deep") }
