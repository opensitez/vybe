// vybe-test: go/lang_interfaces_embedding/interface_satisfied_by_pointer
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs

package main
import "fmt"
type I interface { M() int }
type T int
func (t *T) M() int { return int(*t) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var i I = (*T)(nil)
__check(fmt.Sprint(i == nil), "true") }
