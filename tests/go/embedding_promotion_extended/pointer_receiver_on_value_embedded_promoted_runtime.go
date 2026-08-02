// vybe-test: go/embedding_promotion_extended/pointer_receiver_on_value_embedded_promoted_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type inner struct { n int }
func (i *inner) double() { i.n *= 2 }
type outer struct { inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { o := outer{inner: inner{n: 3}}
o.double()
__check(fmt.Sprint(o.n), "6") }
