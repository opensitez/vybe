// vybe-test: go/interface_embedding_methods/four_level_interface_embed_dispatch_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type d interface { n() int }
type c interface { d }
type b interface { c }
type a interface { b }
type leaf struct { value int }
func (l leaf) n() int { return l.value }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var top a = leaf{value: 13}
__check(fmt.Sprint(top.n()), "13") }
