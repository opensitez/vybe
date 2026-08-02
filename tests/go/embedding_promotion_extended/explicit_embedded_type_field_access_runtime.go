// vybe-test: go/embedding_promotion_extended/explicit_embedded_type_field_access_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type inner struct { name string }
type outer struct { inner
name string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { o := outer{inner: inner{name: "in"}, name: "out"}
__check(fmt.Sprint(o.inner.name), "in") }
