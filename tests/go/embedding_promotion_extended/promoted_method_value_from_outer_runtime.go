// vybe-test: go/embedding_promotion_extended/promoted_method_value_from_outer_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type inner struct { n int }
func (i inner) total() int { return i.n }
type outer struct { inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { o := outer{inner: inner{n: 6}}
fn := o.total
__check(fmt.Sprint(fn()), "6") }
