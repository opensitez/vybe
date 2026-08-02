// vybe-test: go/embedding_promotion_extended/promoted_method_with_args_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type inner struct { base int }
func (i inner) add(d int) int { return i.base + d }
type outer struct { inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { o := outer{inner: inner{base: 3}}
__check(fmt.Sprint(o.add(4)), "7") }
