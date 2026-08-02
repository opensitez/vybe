// vybe-test: go/method_sets_pointer_value/embedded_pointer_field_method_promoted_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type inner struct { tag string }
func (i inner) label() string { return i.tag }
type outer struct { *inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { o := outer{inner: &inner{tag: "go"}}
__check(fmt.Sprint(o.label()), "go") }
