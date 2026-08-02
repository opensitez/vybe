// vybe-test: go/lang_interfaces_embedding/embedding_promotes_method
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs

package main
import "fmt"
type A struct{}
func (A) Hi() string { return "hi" }
type B struct { A }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var b B
__check(fmt.Sprint(b.Hi()), "hi") }
