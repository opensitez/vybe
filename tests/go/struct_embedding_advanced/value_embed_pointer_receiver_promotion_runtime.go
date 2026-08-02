// vybe-test: go/struct_embedding_advanced/value_embed_pointer_receiver_promotion_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

package main
import "fmt"
type inner struct { n int }
func (i *inner) double() { i.n = i.n * 2 }
type outer struct { inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := outer{inner: inner{n: 4}}
value.double()
__check(fmt.Sprint(value.n), "8")
}
