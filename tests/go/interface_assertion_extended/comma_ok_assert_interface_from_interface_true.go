// vybe-test: go/interface_assertion_extended/comma_ok_assert_interface_from_interface_true
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type reader interface { read() int }
type book struct { pages int }
func (b book) read() int { return b.pages }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var concrete reader = book{pages: 10}
var v interface{} = concrete
r, ok := v.(reader)
__check(fmt.Sprint(r.read()), "10")
__check(fmt.Sprint(ok), "true") }
