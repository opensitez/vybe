// vybe-test: go/embedding_promotion_extended/embedded_slice_field_len_promoted_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type inner struct { items []int }
type outer struct { inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { o := outer{inner: inner{items: []int{1, 2, 3}}}
__check(fmt.Sprint(len(o.items)), "3") }
