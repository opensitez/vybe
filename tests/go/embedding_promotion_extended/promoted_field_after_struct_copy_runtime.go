// vybe-test: go/embedding_promotion_extended/promoted_field_after_struct_copy_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type inner struct { n int }
type outer struct { inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := outer{inner: inner{n: 2}}
b := a
b.n = 5
__check(fmt.Sprint(a.n), "2")
__check(fmt.Sprint(b.n), "5") }
