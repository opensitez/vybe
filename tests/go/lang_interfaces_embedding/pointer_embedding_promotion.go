// vybe-test: go/lang_interfaces_embedding/pointer_embedding_promotion
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs

package main
import "fmt"
type A struct{}
func (A) X() int { return 1 }
type B struct { *A }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b := B{A: &A{}}
__check(fmt.Sprint(b.X()), "1") }
