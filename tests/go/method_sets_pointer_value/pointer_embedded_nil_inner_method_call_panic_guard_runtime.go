// vybe-test: go/method_sets_pointer_value/pointer_embedded_nil_inner_method_call_panic_guard_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type inner struct { n int }
func (i *inner) peek() int { if i == nil { return -1 }
return i.n }
type outer struct { *inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var o outer
__check(fmt.Sprint(o.peek()), "-1") }
