// vybe-test: go/lang_interfaces_embedding/type_assertion_to_concrete
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs

package main
import "fmt"
type I interface { M() }
type T struct{}
func (T) M() {}
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var i I = T{}
__check(fmt.Sprint(i.(T) == T{}), "true") }
