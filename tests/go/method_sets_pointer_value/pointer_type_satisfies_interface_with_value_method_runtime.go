// vybe-test: go/method_sets_pointer_value/pointer_type_satisfies_interface_with_value_method_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

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

func main() { var r reader = &book{pages: 120}
__check(fmt.Sprint(r.read()), "120") }
