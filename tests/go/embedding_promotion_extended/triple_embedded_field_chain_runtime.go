// vybe-test: go/embedding_promotion_extended/triple_embedded_field_chain_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type c struct { n int }
type b struct { c }
type a struct { b }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { v := a{b: b{c: c{n: 11}}}
__check(fmt.Sprint(v.n), "11") }
