// vybe-test: go/embedding_promotion_extended/explicit_embedded_type_method_call_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type inner struct{}
func (inner) label() string { return "inner" }
type outer struct { inner }
func (outer) label() string { return "outer" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { o := outer{}
__check(fmt.Sprint(o.inner.label()), "inner") }
