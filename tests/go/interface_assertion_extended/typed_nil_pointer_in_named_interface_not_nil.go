// vybe-test: go/interface_assertion_extended/typed_nil_pointer_in_named_interface_not_nil
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type holder interface { size() int }
type box struct { n int }
func (b *box) size() int { return b.n }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var p *box
var h holder = p
__check(fmt.Sprint(h == nil), "false") }
