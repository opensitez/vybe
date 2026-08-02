// vybe-test: go/embedding_promotion_extended/embedded_func_field_promoted_call_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type inner struct { fn func(int) int }
type outer struct { inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { o := outer{inner: inner{fn: func(x int) int { return x * 2 }}}
__check(fmt.Sprint(o.fn(5)), "10") }
