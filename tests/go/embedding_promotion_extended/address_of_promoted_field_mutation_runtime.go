// vybe-test: go/embedding_promotion_extended/address_of_promoted_field_mutation_runtime
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

func main() { o := outer{inner: inner{n: 1}}
ptr := &o.n
*ptr = 6
__check(fmt.Sprint(o.n), "6") }
