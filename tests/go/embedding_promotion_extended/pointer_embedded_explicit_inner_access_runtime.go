// vybe-test: go/embedding_promotion_extended/pointer_embedded_explicit_inner_access_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type inner struct { count int }
type outer struct { *inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { o := outer{inner: &inner{count: 4}}
__check(fmt.Sprint(o.inner.count), "4") }
