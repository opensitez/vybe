// vybe-test: go/lang_interfaces_embedding/override_promoted_method
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs

package main
import "fmt"
type A struct{}
func (A) Name() string { return "A" }
type B struct { A }
func (B) Name() string { return "B" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(B{}.Name()), "B") }
