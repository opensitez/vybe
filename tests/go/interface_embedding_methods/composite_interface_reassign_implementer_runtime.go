// vybe-test: go/interface_embedding_methods/composite_interface_reassign_implementer_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type sized interface { size() int }
type measurable interface { sized }
type small struct{}
func (small) size() int { return 1 }
type large struct{}
func (large) size() int { return 9 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var m measurable = small{}
__check(fmt.Sprint(m.size()), "1")
m = large{}
__check(fmt.Sprint(m.size()), "9") }
