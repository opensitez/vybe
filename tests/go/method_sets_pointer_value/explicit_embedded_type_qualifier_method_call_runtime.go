// vybe-test: go/method_sets_pointer_value/explicit_embedded_type_qualifier_method_call_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type base struct{}
func (base) tag() string { return "base" }
type derived struct { base }
func (derived) tag() string { return "derived" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { d := derived{}
__check(fmt.Sprint(d.base.tag()), "base") }
